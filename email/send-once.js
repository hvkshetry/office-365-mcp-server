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
 */

const fs = require('fs');
const path = require('path');

const SEND_ONCE_RULING_ID = 'HOUSE-OFFICE-MCP-SEND-ONCE-20261119';  // for HOUSE-TWIN-EMAILS-EXACTLY-ONCE-20261119
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

function stampSendKey(message, key) {
  message.singleValueExtendedProperties = [
    ...(message.singleValueExtendedProperties || []),
    { id: SEND_KEY_PROPERTY, value: key }
  ];
  message.internetMessageHeaders = [
    ...(message.internetMessageHeaders || []),
    { name: SEND_KEY_HEADER, value: key }
  ];
  return message;
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
      writeLedger(once.ledgerPath, recovered);
      return { sent: true, witness: 'sent_items', record: recovered };
    }
    return { sent: false };
  } catch (error) {
    return { sent: false, queryError: error.message };
  }
}

function recordSent(once, details) {
  writeLedger(once.ledgerPath, { key: once.key, job_id: once.jobId, sent_at: new Date().toISOString(),
                                 witness: 'send', ruling: SEND_ONCE_RULING_ID, ...details });
}

module.exports = {
  SEND_ONCE_RULING_ID,
  SEND_KEY_PROPERTY,
  SEND_KEY_HEADER,
  sendOnceOf,
  alreadySent,
  stampSendKey,
  recordSent,
  readLedger
};
