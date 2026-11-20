/**
 * Exactly-once sends for the twin emails (HOUSE-OFFICE-MCP-SEND-ONCE-20261119, for HOUSE-TWIN-EMAILS-EXACTLY-ONCE-20261119).
 *
 * A twin email's dispatch carries `policy.send_once = {job_id, key, ledger_path}` (communication-agent
 * src/policy.py build_execution_policy, from src/send_claims.py). Before a send, office-mcp asks two witnesses whether
 * this key already went out:
 *   1. the sent ledger at `ledger_path`, written right after Graph accepted a send under the key;
 *   2. the mailbox's Sent Items, where every keyed send carries the key as an extended property (filterable) and an
 *      X-Puranos-Send-Key header. This covers a process killed after Graph accepted the send and before it wrote the
 *      ledger.
 * Either one holding the key means the send is answered "already sent" and nothing goes out; else the message is
 * stamped with the key, sent, and the ledger written. A resumed or retried dispatch therefore completes a send that
 * never happened and repeats none. A Sent Items query that fails does not block the send (a missing daily email is
 * the worse failure); the ledger still stops any repeat that it saw.
 *
 * Two failures that must never cost the day's email or fail a send that went out (lead review 2026-11-19):
 *   - Graph REJECTS the stamped sendMail with a definite request-validation error (HTTP 400, `definiteStampRejection`):
 *     nothing went out, so the caller sends once more, unstamped, and the ledger and audit say `stamp_rejected`. A
 *     timeout, a 5xx or any error without that status is ambiguous (the first send may have gone out): never retried.
 *   - the ledger cannot be written after Graph accepted the send: `recordSent` returns the error instead of throwing,
 *     the tool call still succeeds and the audit says `ledger_write_failed`. Sent Items remains the witness.
 * Sent Items is eventually consistent: a message Graph just accepted can take seconds to show there, so a retry
 * within seconds of a kill (after the accept, before the ledger) can miss it. The ledger and the communication-agent's
 * audit row are the witnesses that close that window when they landed.
 *
 * A DRAFT THAT IS THEN SENT is a keyed send too (HOUSE-OFFICE-MCP-KEYED-SEND-COVERS-SENT-DRAFT-20261120): the same
 * ledger, the same Sent Items witness, the same stamp, the same two fallbacks. The key attaches twice:
 *   - when the draft is WRITTEN (`mail draft`): the witnesses are asked first, and the draft is created stamped
 *     (the extended property and the header). Already sent: no draft is written. A stamp Graph refuses with a definite
 *     400 leaves the draft written unstamped (`stamp_rejected`).
 *   - when the draft is SENT (`mail send_draft`): the witnesses are asked again, then the extended property is put on
 *     the draft by PATCH (`draftStampPatch`; the same value where it is already there). That covers a draft written
 *     by any route: a reply-all draft, one written before the dispatch, one whose stamp was refused at creation. The
 *     stamp step sends nothing, so whatever way it fails the draft is still sent, once, and the ledger and audit say
 *     `stamp_rejected`. Then the draft is sent and the ledger written.
 * What a repeat does: `draft` again before anything went out writes another stamped draft (a draft is not an email;
 * the earlier one stays in Drafts); `draft` or `send_draft` after the email went out is answered "already sent" by
 * the ledger, or by Sent Items when the process died after Graph accepted the send and before the ledger.
 */

const fs = require('fs');
const path = require('path');

const SEND_ONCE_RULING_ID = 'HOUSE-OFFICE-MCP-SEND-ONCE-20261119';  // for HOUSE-TWIN-EMAILS-EXACTLY-ONCE-20261119
const KEYED_DRAFT_RULING_ID = 'HOUSE-OFFICE-MCP-KEYED-SEND-COVERS-SENT-DRAFT-20261120';
// A named MAPI property in a namespace of our own: kept on the Sent Items copy and filterable through Graph.
const SEND_KEY_PROPERTY = 'String {6f2c8a51-2b4d-4f7e-9d10-5a0c3e7b2a91} Name PuranosSendKey';
const SEND_KEY_HEADER = 'X-Puranos-Send-Key';

function sendOnceOf(policy) {
  const once = policy && policy.send_once;
  if (!once || typeof once !== 'object' || !once.key || !once.ledger_path) {
    return null;
  }
  return { jobId: String(once.job_id || ''), key: `${once.job_id || ''}/${once.key}`, ledgerPath: String(once.ledger_path) };
}

function readLedger(ledgerPath) {
  try {
    return JSON.parse(fs.readFileSync(ledgerPath, 'utf8'));
  } catch (error) {
    return null;
  }
}

function writeLedger(ledgerPath, record) {
  fs.mkdirSync(path.dirname(ledgerPath), { recursive: true });
  const tmp = `${ledgerPath}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(record));
  fs.renameSync(tmp, ledgerPath);
}

/** A stamped COPY of the message; the caller keeps the unstamped one for the one retry `definiteStampRejection` allows. */
function stampSendKey(message, key) {
  return {
    ...message,
    singleValueExtendedProperties: [
      ...(message.singleValueExtendedProperties || []),
      { id: SEND_KEY_PROPERTY, value: key }
    ],
    internetMessageHeaders: [
      ...(message.internetMessageHeaders || []),
      { name: SEND_KEY_HEADER, value: key }
    ]
  };
}

/**
 * The PATCH body that puts the send key on a draft already written: the filterable extended property, which is what
 * the Sent Items witness reads. (The header is written where the draft is created keyed; it is not needed here.)
 */
function draftStampPatch(key) {
  return { singleValueExtendedProperties: [{ id: SEND_KEY_PROPERTY, value: key }] };
}

/**
 * True only when Graph answered the stamped sendMail with HTTP 400: a request-validation refusal, so nothing was
 * sent. Anything else (a timeout or network error, which carry no status; 429; any 5xx; any other 4xx) is not a
 * stamp rejection: it is either ambiguous or would fail unstamped too.
 */
function definiteStampRejection(error) {
  return Boolean(error) && error.statusCode === 400;
}

function sentItemsQuery(mailboxPrefix, key) {
  const literal = key.replace(/'/g, "''");
  const filter = `singleValueExtendedProperties/Any(ep: ep/id eq '${SEND_KEY_PROPERTY}' and ep/value eq '${literal}')`;
  return {
    endpoint: `${mailboxPrefix}/mailFolders/sentitems/messages`,
    query: { $filter: filter, $select: 'id,subject,sentDateTime', $top: 1 }
  };
}

/**
 * Whether the keyed send already went out: {sent: true, witness, record} or {sent: false, queryError?}.
 * `callGraph(method, endpoint, body, query)` is the caller's Graph function (injected for tests).
 */
async function alreadySent(once, mailboxPrefix, callGraph) {
  const record = readLedger(once.ledgerPath);
  if (record && record.key === once.key) {
    return { sent: true, witness: 'ledger', record };
  }
  try {
    const { endpoint, query } = sentItemsQuery(mailboxPrefix, once.key);
    const found = await callGraph('GET', endpoint, null, query);
    const hit = found && Array.isArray(found.value) && found.value[0];
    if (hit) {
      const recovered = { key: once.key, job_id: once.jobId, message_id: hit.id, subject: hit.subject,
                          sent_at: hit.sentDateTime, witness: 'sent_items', ruling: SEND_ONCE_RULING_ID };
      let ledgerWriteError = null;
      try {
        writeLedger(once.ledgerPath, recovered);
      } catch (error) {   // Sent Items holds the key: the send is still "already sent", whatever the ledger does
        ledgerWriteError = error.message;
      }
      return { sent: true, witness: 'sent_items', record: recovered, ...(ledgerWriteError ? { ledgerWriteError } : {}) };
    }
    return { sent: false };
  } catch (error) {
    return { sent: false, queryError: error.message };
  }
}

/** Write the sent ledger after Graph accepted the send. Never throws: returns null, or the error text when it could not. */
function recordSent(once, details) {
  try {
    writeLedger(once.ledgerPath, { key: once.key, job_id: once.jobId, sent_at: new Date().toISOString(),
                                   witness: 'send', ruling: SEND_ONCE_RULING_ID, ...details });
    return null;
  } catch (error) {
    return error.message;
  }
}

module.exports = {
  SEND_ONCE_RULING_ID,
  KEYED_DRAFT_RULING_ID,
  SEND_KEY_PROPERTY,
  SEND_KEY_HEADER,
  sendOnceOf,
  alreadySent,
  stampSendKey,
  draftStampPatch,
  definiteStampRejection,
  recordSent,
  readLedger
};
