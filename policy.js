/**
 * Runtime side-effect policy and audit hooks for Office MCP.
 *
 * Policy is injected per communication-agent run through environment variables
 * on the Office MCP process in that agent's ephemeral `.mcp.json`.
 */

let _cachedPolicy;

const LEGAL_MAILBOX = 'legal.agent@circleh2o.com';
const LEGAL_TOOL_MANIFEST = Object.freeze({
  system: new Set(['about', 'check_status']),
  mail: new Set([
    'list', 'read', 'get_attachment', 'cleanup_attachment', 'search', 'focused',
    'list_drafts', 'draft', 'update_draft', 'draft_reply_all'
  ]),
  calendar: new Set(['list', 'get']),
  teams_meeting: new Set([
    'get', 'find_by_url', 'list_transcripts', 'get_transcript',
    'list_recordings', 'get_recording', 'get_participants', 'get_insights'
  ]),
  teams_channel: new Set(['list', 'get', 'list_messages', 'get_message', 'list_members', 'list_tabs']),
  teams_chat: new Set(['list', 'get', 'list_messages', 'get_message', 'list_members']),
  planner: new Set(['list', 'get', 'get_details', 'get_assignments', 'get_tasks', 'lookup']),
  files: new Set(['list', 'get', 'download', 'search']),
  files_map_sharepoint_path: null,
  search: null,
  contacts: new Set(['list', 'search', 'get', 'list_folders']),
  notifications: new Set(['list']),
  todo: new Set(['list_lists', 'get_list', 'list_tasks', 'get_task', 'list_checklist']),
  groups: new Set(['list', 'get', 'list_members', 'list_owners', 'get_drive', 'list_drives']),
  directory: new Set(['lookup_user', 'get_profile', 'get_manager', 'get_reports', 'get_presence', 'search_users'])
});

function isLegalDraftOnly() {
  return ['1', 'true', 'yes'].includes(String(process.env.LEGAL_AGENT_DRAFT_ONLY || '').toLowerCase());
}

function getExecutionPolicy() {
  if (_cachedPolicy !== undefined) {
    return _cachedPolicy;
  }

  const raw = process.env.COMMUNICATION_AGENT_POLICY_JSON;
  if (!raw) {
    _cachedPolicy = null;
    return _cachedPolicy;
  }

  try {
    _cachedPolicy = JSON.parse(raw);
  } catch (error) {
    console.error('[POLICY] Failed to parse COMMUNICATION_AGENT_POLICY_JSON:', error.message);
    _cachedPolicy = { invalid: true };
  }

  return _cachedPolicy;
}

async function recordSideEffect(effectType, status, target = null, details = null) {
  const url = process.env.COMMUNICATION_AGENT_AUDIT_URL;
  const workItemId = process.env.COMMUNICATION_AGENT_WORK_ITEM_ID;
  if (!url || !workItemId || typeof fetch !== 'function') {
    return;
  }

  const headers = {
    'Content-Type': 'application/json'
  };
  if (process.env.COMMUNICATION_AGENT_AUDIT_TOKEN) {
    headers['X-Internal-Token'] = process.env.COMMUNICATION_AGENT_AUDIT_TOKEN;
  }

  try {
    await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        work_item_id: workItemId,
        effect_type: effectType,
        status,
        target,
        details
      })
    });
  } catch (error) {
    console.error('[POLICY] Failed to audit side effect:', error.message);
  }
}

function _mailOperationAllowed(policy, operation) {
  const mailPolicy = policy?.mail;
  if (!mailPolicy) {
    return true;
  }

  switch (operation) {
    case 'draft':
      return !!mailPolicy.draft;
    case 'draft_reply_all':
      return !!mailPolicy.draft_reply_all;
    case 'update_draft':
      return !!mailPolicy.update_draft;
    case 'send':
      return !!mailPolicy.send;
    case 'reply':
      return !!mailPolicy.reply;
    case 'send_draft':
      return !!mailPolicy.send_draft;
    default:
      return false;
  }
}

// --------------------------------------------------------------------------
// Owner-identity block (owner D10, 2026-11-16,
// HOUSE-OCI-AGENTS-NEVER-SEND-AS-OWNER-20261116, superseding
// HOUSE-OCI-AGENTS-NEVER-ACT-AS-OWNER-20261116): a deployment switch, never a
// per-call choice. Set (1/true/yes), this process refuses what would send or
// post as the signed-in owner: mail send/reply/send_draft and create_rule from
// the owner's own mailbox (no `mailbox`, `me`, an owner address, or a value
// that is not one address); every calendar create/update/delete and Teams
// meeting create/update/cancel (both act only from the owner's account and can
// send invitations, updates or cancellations); every Teams channel or chat
// message post, edit or delete. Drafts in the owner's mailbox (draft,
// draft_reply_all, update_draft) send nothing and stay allowed; reads are
// untouched. Unset (default) keeps the owner's local behaviour; the OCI
// communication agent sets it on every persona it dispatches. Owner addresses:
// OFFICE_MCP_OWNER_MAILBOXES (comma list), else DEFAULT_OWNER_MAILBOXES.
// --------------------------------------------------------------------------
const OWNER_IDENTITY_SWITCH = 'OFFICE_MCP_BLOCK_OWNER_IDENTITY';
const OWNER_IDENTITY_REFUSAL = 'OWNER_IDENTITY_BLOCKED';
const DEFAULT_OWNER_MAILBOXES = Object.freeze(['hersh@circleh2o.com']);
const OWNER_MAIL_WRITE_OPS = new Set(['send', 'reply', 'send_draft', 'create_rule']);
// No mailbox parameter: these act only as the owner.
const OWNER_ACCOUNT_WRITE_OPS = Object.freeze({
  calendar: new Set(['create', 'update', 'delete']),
  teams_meeting: new Set(['create', 'update', 'cancel'])
});
// teams_chat maps the legacy send/update/delete onto the message operations.
const OWNER_TEAMS_MESSAGE_OPS = Object.freeze({
  teams_channel: new Set(['create_message', 'reply_to_message']),
  teams_chat: new Set([
    'send_message', 'update_message', 'delete_message', 'send', 'update', 'delete'
  ])
});

function isOwnerIdentityBlocked() {
  return ['1', 'true', 'yes'].includes(
    String(process.env[OWNER_IDENTITY_SWITCH] || '').trim().toLowerCase()
  );
}

function ownerMailboxes() {
  const raw = process.env.OFFICE_MCP_OWNER_MAILBOXES;
  const list = raw ? raw.split(/[,;\s]+/) : DEFAULT_OWNER_MAILBOXES;
  return new Set(list.map(s => s.trim().toLowerCase()).filter(Boolean));
}

function _isSingleAddress(value) {
  const parts = value.split('@');
  return parts.length === 2 && parts[0] !== '' && parts[1] !== '' && !/\s/.test(value);
}

// As teams_chat infers an operation-less legacy call.
function _teamsChatOperation(args) {
  if (args?.operation) return args.operation;
  if (args?.chatId && args?.messageId) return 'get_message';
  if (args?.chatId && args?.content) return 'send_message';
  return undefined;
}

function ownerIdentityRefusal(toolName, args = {}) {
  if (!isOwnerIdentityBlocked()) {
    return null;
  }
  if (toolName === 'mail') {
    const operation = args?.operation;
    if (!OWNER_MAIL_WRITE_OPS.has(operation)) {
      return null;
    }
    const mailbox = String(args.mailbox ?? '').trim().toLowerCase();
    if (mailbox && mailbox !== 'me' && _isSingleAddress(mailbox) && !ownerMailboxes().has(mailbox)) {
      return null;
    }
    const named = mailbox ? `mailbox "${String(args.mailbox).trim()}"` : 'no mailbox (the owner\'s own)';
    return `${OWNER_IDENTITY_REFUSAL}: mail ${operation} with ${named} would act as the signed-in owner; `
      + `this deployment (${OWNER_IDENTITY_SWITCH}) acts only from an agent mailbox: pass `
      + 'mailbox: "<your agent mailbox address>"';
  }
  const accountOps = OWNER_ACCOUNT_WRITE_OPS[toolName];
  if (accountOps && accountOps.has(args?.operation)) {
    return `${OWNER_IDENTITY_REFUSAL}: ${toolName} ${args.operation} acts from the signed-in owner's account `
      + `and can send invitations, updates or cancellations as the owner; this deployment `
      + `(${OWNER_IDENTITY_SWITCH}) refuses it`;
  }
  const messageOps = OWNER_TEAMS_MESSAGE_OPS[toolName];
  if (messageOps) {
    const operation = toolName === 'teams_chat' ? _teamsChatOperation(args) : args?.operation;
    if (messageOps.has(operation)) {
      return `${OWNER_IDENTITY_REFUSAL}: ${toolName} ${operation} posts as the signed-in owner; `
        + `this deployment (${OWNER_IDENTITY_SWITCH}) refuses every Teams message write made as the owner`;
    }
  }
  return null;
}

async function enforceOfficeToolPolicy(toolName, args = {}) {
  const ownerRefusal = ownerIdentityRefusal(toolName, args);
  if (ownerRefusal) {
    await recordSideEffect(`${toolName}.${args?.operation || 'unknown'}`, 'blocked', args?.mailbox || null, {
      reason: ownerRefusal
    });
    return ownerRefusal;
  }
  const idRefusal = applyRequestEmailId(toolName, args);
  if (idRefusal) {
    await recordSideEffect(`${toolName}.${args?.operation || 'unknown'}`, 'blocked', args?.mailbox || null, {
      reason: idRefusal
    });
    return idRefusal;
  }
  const forwardRefusal = await recipientGateRefusal('create_rule', () => args?.forwardTo, args, {
    applies: toolName === 'mail' && args?.operation === 'create_rule' && _isPassed(args?.forwardTo)
  });
  if (forwardRefusal) {
    return forwardRefusal;
  }
  if (!isLegalDraftOnly()) {
    return null;
  }

  if (!Object.prototype.hasOwnProperty.call(LEGAL_TOOL_MANIFEST, toolName)) {
    const reason = `Blocked by Legal OS tool manifest: ${toolName} is not reviewed`;
    await recordSideEffect('legal.tool.blocked', 'blocked', toolName, { reason });
    return reason;
  }

  const allowedOperations = LEGAL_TOOL_MANIFEST[toolName];
  const operation = args?.operation;
  // R6: under a request's recipient gate Legal transmits to the promoter group (the mail handler checks every
  // recipient); without one Legal stays draft-only.
  const gatedTransmit = toolName === 'mail' && LEGAL_GATED_TRANSMIT_OPERATIONS.has(operation) && recipientGate() !== null;
  if (allowedOperations !== null && !gatedTransmit && (!operation || !allowedOperations.has(operation))) {
    const reason = `Blocked by Legal OS tool manifest: ${toolName}.${operation || '<missing>'} is not allowed`;
    await recordSideEffect('legal.tool.blocked', 'blocked', toolName, { reason, operation });
    return reason;
  }

  if (toolName === 'mail' && (gatedTransmit || ['draft', 'update_draft', 'draft_reply_all'].includes(operation))) {
    const mailbox = String(args.mailbox || '').toLowerCase();
    if (mailbox !== LEGAL_MAILBOX) {
      const reason = `Blocked by Legal OS: ${operation} must explicitly use mailbox "${LEGAL_MAILBOX}"`;
      await recordSideEffect(`mail.${operation}`, 'blocked', mailbox || null, { reason });
      return reason;
    }
  }
  return null;
}

// --------------------------------------------------------------------------
// Recipient-domain enforcement.
//
// OPT-IN: active only when a policy carries `allowed_recipient_domains`. No
// persona sets an execution policy today, so this is inert for all of them
// until one opts in.
//
// It lives here rather than in a Claude Code PreToolUse hook because the
// production harness on purposeenergy-base is Codex (DEFAULT_AI_BACKEND=codex,
// every persona pinned to it), where Claude hooks never fire and
// approval_policy is "never". A hook-based fence does not exist on the path
// that actually runs; a check in the MCP server holds under any harness.
// --------------------------------------------------------------------------

// Pull addresses out of every shape the mail tool accepts: a bare string, a
// comma/semicolon-separated string, an array, "Name <a@b.com>", or a
// Graph-style {emailAddress:{address}}. Anything unrecognised yields a
// sentinel that fails the domain test, so an unexpected shape denies rather
// than slipping through.
function _collectAddresses(value, out) {
  if (value === null || value === undefined) return;
  if (typeof value === 'string') {
    value.split(/[,;]/).map(s => s.trim()).filter(Boolean).forEach(s => out.push(s));
    return;
  }
  if (Array.isArray(value)) {
    value.forEach(v => _collectAddresses(v, out));
    return;
  }
  if (typeof value === 'object') {
    const direct = value.address || value.email || value.Address;
    const nested = value.emailAddress || value.EmailAddress;
    if (direct) {
      out.push(String(direct));
    } else if (nested && typeof nested === 'object' && (nested.address || nested.Address)) {
      out.push(String(nested.address || nested.Address));
    } else {
      out.push('<unparseable-recipient>');
    }
    return;
  }
  out.push('<unparseable-recipient>');
}

// Exact domain comparison. A suffix test would accept
// user@purposeenergy.com.evil.co.
function _domainOf(raw) {
  let s = String(raw).trim();
  const angled = s.match(/<([^<>]+)>/);
  if (angled) s = angled[1].trim();
  s = s.replace(/^</, '').replace(/>$/, '').trim();
  if ((s.match(/@/g) || []).length !== 1) return '';
  return s.split('@')[1].trim().replace(/\.$/, '').toLowerCase();
}

function _recipientViolation(policy, params) {
  const allowed = policy.allowed_recipient_domains;
  if (!Array.isArray(allowed) || allowed.length === 0) return null;
  const allowSet = new Set(allowed.map(d => String(d).trim().toLowerCase()));

  const addrs = [];
  ['to', 'cc', 'bcc', 'toRecipients', 'ccRecipients', 'bccRecipients']
    .forEach(f => _collectAddresses(params[f], addrs));
  if (addrs.length === 0) return null;   // nothing addressed in this call

  const external = addrs.filter(a => !allowSet.has(_domainOf(a)));
  if (external.length === 0) return null;
  return [...new Set(external)].slice(0, 5).join(', ');
}

// --------------------------------------------------------------------------
// Recipient gate (owner rulings 2026-11-18: HOUSE-CBG-ORIGINATION-PROMOTER-THREAD-REPLY-20261118 for CBG Origination,
// HOUSE-LEGAL-PROMOTER-GROUP-REPLY-20261118 for Legal). Active only when the per-request policy
// (COMMUNICATION_AGENT_POLICY_JSON) carries `recipient_gate`; the communication agent sets it for the gated personas
// (communication-agent src/policy.py RECIPIENT_GATE_RULINGS), and its delegates inherit the request policy. Under it a
// transmission (mail send, reply, send_draft, or a create_rule that forwards) is allowed only when EVERY recipient is
// one of `recipient_gate.allowed_recipients`: to, cc and bcc, a draft's own recipients, and for a reply the set Graph
// derives from the original message (reply: from/sender/replyTo; reply-all adds its To and Cc; the replying mailbox
// itself excluded, as Graph excludes it). Anything else is refused and the persona drafts (draft_reply_all on a thread,
// draft otherwise) for the owner to send. Drafts send nothing and are not gated. A gate with no readable allowed list,
// or a transmission whose recipients cannot be read, is refused (fail closed). Under a gate the Legal OS draft-only
// switch admits the gated transmissions (R6); without one Legal stays draft-only.
// --------------------------------------------------------------------------
const RECIPIENT_GATE_REFUSAL = 'RECIPIENT_GATE';
const GATED_TRANSMIT_OPERATIONS = Object.freeze(new Set(['send', 'reply', 'send_draft']));
const LEGAL_GATED_TRANSMIT_OPERATIONS = GATED_TRANSMIT_OPERATIONS;

function _isPassed(value) {
  if (value === undefined || value === null || value === '') return false;
  return !(Array.isArray(value) && value.length === 0);
}

// One bare, lower-cased address, or '' when the value is not exactly one address.
function _bareAddress(raw) {
  let s = String(raw ?? '').trim();
  const angled = s.match(/<([^<>]+)>/);
  if (angled) s = angled[1].trim();
  s = s.replace(/^mailto:/i, '').trim().toLowerCase();
  return _isSingleAddress(s) ? s : '';
}

function recipientGate(policy = getExecutionPolicy()) {
  const gate = policy?.recipient_gate;
  if (gate === undefined || gate === null) return null;
  const listed = Array.isArray(gate?.allowed_recipients) ? gate.allowed_recipients : [];
  return {
    allowed: new Set(listed.map(_bareAddress).filter(Boolean)),
    ruling: String(gate?.ruling_id || 'the request recipient gate')
  };
}

/**
 * The gate's refusal for one transmission, or null. `recipients` is a value or a function returning one (called only
 * when the gate applies, so an ungated run makes no Graph read); every shape _collectAddresses reads is accepted.
 */
async function recipientGateRefusal(operation, recipients, params = {}, { applies } = {}) {
  const gate = recipientGate();
  const transmits = applies === undefined ? GATED_TRANSMIT_OPERATIONS.has(operation) : applies;
  if (!gate || !transmits) return null;
  let read;
  try {
    read = typeof recipients === 'function' ? await recipients() : recipients;
  } catch (error) {
    read = undefined;
    const reason = `${RECIPIENT_GATE_REFUSAL}: mail ${operation} refused: its recipients could not be read `
      + `(${error.message}); under ${gate.ruling} a transmission goes ahead only when every recipient is known to be `
      + 'in the promoter group. Nothing was sent. Draft it instead (draft_reply_all on a thread, else draft) and name '
      + 'the draft.';
    await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, { reason });
    return reason;
  }
  const raw = [];
  _collectAddresses(read, raw);
  const outside = [...new Set(raw.map(a => _bareAddress(a) || String(a)).filter(a => !gate.allowed.has(a)))];
  if (raw.length > 0 && outside.length === 0) return null;
  const why = raw.length === 0
    ? 'names no recipient this check could read'
    : `would go to ${outside.length} recipient(s) outside the promoter group and the agent mailboxes `
      + `(${outside.slice(0, 5).join(', ')})`;
  const instead = operation === 'create_rule'
    ? 'No rule was created: a rule may not forward outside them.'
    : 'Nothing was sent. Answer with mail draft_reply_all on the thread (or draft for a new message), keeping every '
      + 'recipient, and name the draft: the owner sends it.';
  const reason = `${RECIPIENT_GATE_REFUSAL}: mail ${operation} ${why}; under ${gate.ruling} this mailbox transmits `
    + `only to the promoter group (and the agent mailboxes). ${instead}`;
  await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, { reason, outside });
  return reason;
}

// --------------------------------------------------------------------------
// The request's own message id (live test 2026-11-18 02:15 UTC: a woken lead retyped the 152-character Graph id with
// 32 characters missing and Graph answered "Id is malformed"). An email request's policy names the message the run
// answers (`reply_email_id`, communication-agent src/policy.py). mail emailId "original", and a reply or
// draft_reply_all with no emailId, take that id (written into the call's args here, before the handler runs); a reply
// or draft_reply_all naming any other id is refused, naming where the id comes from. Without the key nothing changes.
// --------------------------------------------------------------------------
const ORIGINAL_EMAIL_ID = 'original';
const REQUEST_EMAIL_ID_REFUSAL = 'REQUEST_EMAIL_ID';
const REQUEST_EMAIL_ID_OPERATIONS = Object.freeze(new Set(['reply', 'draft_reply_all']));

function requestEmailId(policy = getExecutionPolicy()) {
  const own = policy?.reply_email_id;
  return typeof own === 'string' && own.trim() ? own.trim() : null;
}

/** Substitutes the request's own id into `args.emailId` in place; returns a refusal or null. */
function applyRequestEmailId(toolName, args) {
  if (toolName !== 'mail' || !args || typeof args !== 'object') return null;
  const own = requestEmailId();
  const given = typeof args.emailId === 'string' ? args.emailId.trim() : args.emailId;
  const named = typeof given === 'string' && given.toLowerCase() === ORIGINAL_EMAIL_ID;
  const answers = REQUEST_EMAIL_ID_OPERATIONS.has(args.operation);
  if (!own) {
    return named
      ? `${REQUEST_EMAIL_ID_REFUSAL}: emailId "original" names the message this run was dispatched to answer, and `
        + 'this run names none (no reply_email_id in COMMUNICATION_AGENT_POLICY_JSON): pass the message id itself.'
      : null;
  }
  if (named || (answers && !_isPassed(given))) {
    args.emailId = own;
    return null;
  }
  if (answers && given !== own) {
    return `${REQUEST_EMAIL_ID_REFUSAL}_MISMATCH: mail ${args.operation} answers the message this run was dispatched `
      + `for, and the emailId given (${String(given ?? '').length} characters) is not its id (${own.length} `
      + 'characters). Nothing was sent or drafted. Pass emailId: "original" and office-mcp uses the request\'s own id '
      + '(COMMUNICATION_AGENT_POLICY_JSON reply_email_id); never retype the id.';
  }
  return null;
}

async function enforceMailPolicy(operation, params = {}) {
  const policy = getExecutionPolicy();
  if (isLegalDraftOnly() && GATED_TRANSMIT_OPERATIONS.has(operation) && !recipientGate(policy)) {
    const reason = `Blocked by Legal OS draft-only policy: mail ${operation} is forbidden`;
    await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, { reason });
    return reason;
  }
  if (isLegalDraftOnly() && ['draft', 'update_draft', 'draft_reply_all'].includes(operation)) {
    const mailbox = String(params.mailbox || '').toLowerCase();
    if (mailbox !== LEGAL_MAILBOX) {
      return `Blocked by Legal OS: ${operation} must explicitly use mailbox "${LEGAL_MAILBOX}"`;
    }
  }
  if (policy?.invalid) {
    return 'Blocked: COMMUNICATION_AGENT_POLICY_JSON is invalid';
  }
  if (!policy) {
    return null;
  }

  // Checked on anything that names recipients, including `draft` and
  // `update_draft`: a draft addressed externally is one click from being sent,
  // and update_draft is precisely how an internally-addressed draft gets
  // re-pointed at an outsider.
  if (['send', 'reply', 'draft', 'update_draft', 'draft_reply_all'].includes(operation)) {
    const external = _recipientViolation(policy, params);
    if (external) {
      const reason =
        `Blocked by communication-agent policy: mail ${operation} names external `
        + `recipient(s) ${external}; this agent may address only `
        + `${policy.allowed_recipient_domains.join(', ')}`;
      await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, {
        reason,
        external,
        allowedDomains: policy.allowed_recipient_domains
      });
      return reason;
    }
  }

  if (!_mailOperationAllowed(policy, operation)) {
    const reason = `Blocked by communication-agent policy: mail ${operation} is not allowed`;
    await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, {
      reason,
      replyActor: policy.reply_actor || null
    });
    return reason;
  }

  if (['send', 'reply', 'send_draft'].includes(operation) && policy.authorized_mailbox) {
    const mailbox = params.mailbox || 'me';
    if (mailbox.toLowerCase() !== policy.authorized_mailbox.toLowerCase()) {
      const reason = (
        `Blocked by communication-agent policy: external ${operation} must use mailbox `
        + `"${policy.authorized_mailbox}"`
      );
      await recordSideEffect(`mail.${operation}`, 'blocked', mailbox, {
        reason,
        replyActor: policy.reply_actor || null,
        authorizedMailbox: policy.authorized_mailbox
      });
      return reason;
    }
  }

  return null;
}

async function enforcePlannerPolicy(args = {}) {
  const policy = getExecutionPolicy();
  if (!policy?.planner) {
    return null;
  }

  const readOnlyOps = new Set(['list', 'get', 'get_details', 'get_assignments', 'get_tasks', 'lookup']);
  const operation = args.operation || '';
  if (readOnlyOps.has(operation)) {
    return null;
  }

  if (!policy.planner.write) {
    const reason = `Blocked by communication-agent policy: planner ${operation} is not allowed`;
    await recordSideEffect('planner.write', 'blocked', args.taskId || args.planId || args.bucketId || null, {
      reason,
      operation,
      entity: args.entity || null
    });
    return reason;
  }

  return null;
}

async function enforceCalendarPolicy(args = {}) {
  const policy = getExecutionPolicy();
  if (!policy?.calendar) {
    return null;
  }

  const operation = args.operation || '';
  if (['list', 'get'].includes(operation)) {
    return null;
  }

  if (!policy.calendar.write) {
    const reason = `Blocked by communication-agent policy: calendar ${operation} is not allowed`;
    await recordSideEffect('calendar.write', 'blocked', args.eventId || null, {
      reason,
      operation
    });
    return reason;
  }

  return null;
}

module.exports = {
  LEGAL_TOOL_MANIFEST,
  LEGAL_GATED_TRANSMIT_OPERATIONS,
  GATED_TRANSMIT_OPERATIONS,
  RECIPIENT_GATE_REFUSAL,
  REQUEST_EMAIL_ID_REFUSAL,
  ORIGINAL_EMAIL_ID,
  recipientGate,
  recipientGateRefusal,
  requestEmailId,
  applyRequestEmailId,
  OWNER_IDENTITY_SWITCH,
  OWNER_IDENTITY_REFUSAL,
  getExecutionPolicy,
  isLegalDraftOnly,
  isOwnerIdentityBlocked,
  ownerIdentityRefusal,
  recordSideEffect,
  enforceOfficeToolPolicy,
  enforceMailPolicy,
  enforcePlannerPolicy,
  enforceCalendarPolicy
};
