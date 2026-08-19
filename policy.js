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
    // An inbox rule can auto-forward to an outside address, which would carry
    // mail past the recipient fence below without ever calling send/reply.
    // It therefore defaults to DENY under any policy: a persona that genuinely
    // needs it must say so explicitly with `mail.create_rule: true`.
    case 'create_rule':
      return !!mailPolicy.create_rule;
    // Reads, and mailbox hygiene that cannot exfiltrate (move between folders,
    // categories, create_folder), are deliberately not gated here.
    default:
      return false;
  }
}

async function enforceOfficeToolPolicy(toolName, args = {}) {
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
  if (allowedOperations !== null && (!operation || !allowedOperations.has(operation))) {
    const reason = `Blocked by Legal OS tool manifest: ${toolName}.${operation || '<missing>'} is not allowed`;
    await recordSideEffect('legal.tool.blocked', 'blocked', toolName, { reason, operation });
    return reason;
  }

  if (toolName === 'mail' && ['draft', 'update_draft', 'draft_reply_all'].includes(operation)) {
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
// Active when a policy carries `allowed_recipient_domains`,
// `allowed_recipient_addresses`, or both. Every deployed persona now sets one,
// so this is live for all of them -- do not reintroduce the old note claiming
// it is inert. A persona with neither key is unrestricted by recipient.
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

// Strict single-mailbox parser. Returns the canonical lowercase address, or ''
// to DENY.
//
// The permissive version of this took the first `<...>` it found and ignored
// everything around it, so all three of these were approved on the strength of
// the bracketed part while Exchange could plausibly deliver somewhere else:
//
//   attacker@evil.example <inside@purposeenergy.com>
//   Display <inside@purposeenergy.com> <attacker@evil.example>
//   Display <inside@purposeenergy.com> trailing-junk
//
// Ambiguity is rejected rather than resolved: if the fence and the mail system
// could disagree about which mailbox a string names, it does not go out.
function _parseAddress(raw) {
  const s = String(raw).trim();
  if (!s) return '';

  let addr = s;
  const lt = s.indexOf('<');
  const gt = s.indexOf('>');

  if (lt !== -1 || gt !== -1) {
    if (lt === -1 || gt === -1) return '';
    if (s.indexOf('<', lt + 1) !== -1) return '';   // a second '<'
    if (s.indexOf('>', gt + 1) !== -1) return '';   // a second '>'
    if (gt < lt) return '';
    if (gt !== s.length - 1) return '';             // trailing junk after '>'
    const display = s.slice(0, lt);
    if (/[@<>]/.test(display)) return '';           // display name posing as an address
    addr = s.slice(lt + 1, gt).trim();
  }

  addr = addr.replace(/\.$/, '');                   // tolerate one trailing dot
  if (/[\s<>,;]/.test(addr)) return '';
  const parts = addr.split('@');
  if (parts.length !== 2) return '';
  const [local, domain] = parts;
  if (!local || !domain) return '';
  if (domain.startsWith('.') || domain.endsWith('.')) return '';
  if (!domain.includes('.')) return '';
  return addr.toLowerCase();
}

// Both derive from the same parse, so the address check and the domain check can
// never disagree about what a string names.
function _addressOf(raw) {
  return _parseAddress(raw);
}

// A policy may bound recipients by domain, by exact address, or both. Mirrors
// the inbound model in the worker's config.py, where an agent either inherits
// the org-wide domain check or is pinned to an explicit list of people.
// Empty/absent on BOTH fields means no recipient restriction is configured.
function _recipientAllowSets(policy) {
  const clean = (v) => (Array.isArray(v) ? v : [])
    .map(x => String(x).trim().toLowerCase())
    .filter(Boolean);
  const domains = clean(policy.allowed_recipient_domains);
  const addresses = clean(policy.allowed_recipient_addresses);
  if (domains.length === 0 && addresses.length === 0) return null;
  return { domains: new Set(domains), addresses: new Set(addresses) };
}

function _addressAllowed(sets, raw) {
  const addr = _addressOf(raw);
  if (!addr) return false;                          // unparseable or ambiguous
  if (sets.addresses.has(addr)) return true;
  const domain = _domainOf(raw);
  return !!domain && sets.domains.has(domain);
}

// Exact domain comparison, derived from the same strict parse so a suffix trick
// (user@purposeenergy.com.evil.co) or an ambiguous string cannot slip between
// the two checks.
function _domainOf(raw) {
  const addr = _parseAddress(raw);
  return addr ? addr.split('@')[1] : '';
}

function _recipientViolation(policy, params) {
  const sets = _recipientAllowSets(policy);
  if (!sets) return null;

  const addrs = [];
  // `forwardTo` belongs here: an inbox rule that forwards is an outbound path,
  // and it never touches send/reply.
  // Every field here is a way mail or an invitation reaches a person:
  // `forwardTo` is an inbox rule, `attendees` is a calendar invite (Graph mails
  // them automatically), `recipients` is a file-share invitation.
  ['to', 'cc', 'bcc', 'toRecipients', 'ccRecipients', 'bccRecipients',
   'forwardTo', 'attendees', 'recipients']
    .forEach(f => _collectAddresses(params[f], addrs));
  if (addrs.length === 0) return null;   // nothing addressed in this call

  const external = addrs.filter(a => !_addressAllowed(sets, a));
  if (external.length === 0) return null;
  return [...new Set(external)].slice(0, 5).join(', ');
}

// Split an already-resolved Graph recipient list into the ones this persona may
// address and the ones it may not.
//
// This exists because `reply`/`replyAll` name no recipients in their params --
// Graph expands the thread server-side -- so _recipientViolation above sees an
// empty address list and passes. The reply path resolves the real list first
// and calls this, then sends to `kept` explicitly. `enforced:false` means no
// recipient restriction is configured, and the caller must not alter anything.
// Cheap pre-check so the reply path can skip two Graph round-trips when no
// recipient restriction is configured for this persona.
function recipientRestrictionActive() {
  const policy = getExecutionPolicy();
  return !!(policy && _recipientAllowSets(policy));
}

function partitionRecipients(recipients) {
  const policy = getExecutionPolicy();
  const sets = policy ? _recipientAllowSets(policy) : null;
  if (!sets) return { kept: recipients || [], removed: [], enforced: false };

  const kept = [];
  const removed = [];
  for (const r of recipients || []) {
    const addr = r?.emailAddress?.address;
    if (addr && _addressAllowed(sets, addr)) kept.push(r);
    else removed.push(addr || '<unparseable-recipient>');
  }
  return { kept, removed, enforced: true };
}

// Mail operations that must never run unpoliced once enforcement is required.
// Reads are absent on purpose: a missing policy should stop an agent acting,
// not stop it looking.
const _MAIL_ENFORCED_OPS = new Set([
  'send', 'reply', 'draft', 'update_draft', 'send_draft', 'create_rule'
]);

// Phase 3 switch. Until COMMUNICATION_AGENT_POLICY_REQUIRED is truthy, an agent
// with no policy behaves as it always has (unrestricted); once it is set, a
// missing OR malformed policy refuses the operations above instead of silently
// permitting them. Kept as an env flag so enabling it is a config change on a
// verified fleet, not a code change.
function _policyRequired() {
  const raw = String(process.env.COMMUNICATION_AGENT_POLICY_REQUIRED || '').trim().toLowerCase();
  return raw === '1' || raw === 'true' || raw === 'yes';
}

async function enforceMailPolicy(operation, params = {}) {
  const policy = getExecutionPolicy();
  if (isLegalDraftOnly() && ['send', 'reply', 'send_draft'].includes(operation)) {
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
    if (_policyRequired() && _MAIL_ENFORCED_OPS.has(operation)) {
      const reason =
        `Blocked by communication-agent policy: mail ${operation} requires an execution `
        + `policy, and none is configured for this agent. Set COMMUNICATION_AGENT_POLICY_JSON.`;
      await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, {
        reason,
        cause: 'policy_missing_or_malformed'
      });
      return reason;
    }
    return null;
  }

  // Checked on anything that names recipients, including `draft` and
  // `update_draft`: a draft addressed externally is one click from being sent,
  // and update_draft is precisely how an internally-addressed draft gets
  // re-pointed at an outsider.
  if (['send', 'reply', 'draft', 'update_draft', 'draft_reply_all', 'create_rule'].includes(operation)) {
    const external = _recipientViolation(policy, params);
    if (external) {
      // Describe whichever of the two lists is actually configured. Reading
      // allowed_recipient_domains unconditionally throws for an
      // addresses-only policy, which would turn a block into a crash.
      const allowedDesc = [
        ...(policy.allowed_recipient_domains || []).map(d => `@${d}`),
        ...(policy.allowed_recipient_addresses || [])
      ].join(', ') || '(nothing)';
      const reason =
        `Blocked by communication-agent policy: mail ${operation} names external `
        + `recipient(s) ${external}; this agent may address only ${allowedDesc}`;
      await recordSideEffect(`mail.${operation}`, 'blocked', params.mailbox || null, {
        reason,
        external,
        allowedDomains: policy.allowed_recipient_domains || [],
        allowedAddresses: policy.allowed_recipient_addresses || []
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
  if (!policy) {
    return null;
  }

  const operation = args.operation || '';
  if (['list', 'get'].includes(operation)) {
    return null;                                    // reads are never gated
  }

  // Creating an event WITH attendees makes Graph mail an invitation to each of
  // them, so an attendee list is an outbound recipient path that never touches
  // the mail tool. Subject and body ride along with it.
  const external = _recipientViolation(policy, args);
  if (external) {
    const reason =
      `Blocked by communication-agent policy: calendar ${operation} names external `
      + `attendee(s) ${external}; Graph would send them an invitation`;
    await recordSideEffect(`calendar.${operation}`, 'blocked', null, { reason, external });
    return reason;
  }

  // A policy that says nothing about calendar denies writes, rather than
  // silently permitting them the way the old `if (!policy?.calendar)` did.
  if (!policy.calendar) {
    const reason =
      `Blocked by communication-agent policy: calendar ${operation} is not allowed `
      + `(no calendar capability in this agent's policy)`;
    await recordSideEffect(`calendar.${operation}`, 'blocked', null, { reason });
    return reason;
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

// File sharing is an outbound path with two distinct risks: an invitation mails
// a named recipient, and a link-scope share can mint an ANONYMOUS URL that needs
// no identity at all. `shareFile` defaults to scope 'anonymous', so silence here
// is the dangerous default.
async function enforceFilesPolicy(args = {}) {
  const policy = getExecutionPolicy();
  if (!policy) return null;

  const operation = args.operation || '';
  if (operation !== 'share') return null;           // reads/uploads are not gated here

  const files = policy.files || {};
  if (!files.share) {
    const reason =
      'Blocked by communication-agent policy: files share is not allowed for this agent';
    await recordSideEffect('files.share', 'blocked', args.fileId || args.filePath || null, { reason });
    return reason;
  }

  const scope = args.scope || 'anonymous';
  const named = Array.isArray(args.recipients) && args.recipients.length > 0;
  if (!named && !files.allow_anonymous_links) {
    const reason =
      `Blocked by communication-agent policy: a '${scope}' share with no named `
      + `recipients creates a link that anyone holding it can use`;
    await recordSideEffect('files.share', 'blocked', args.fileId || args.filePath || null, {
      reason, scope
    });
    return reason;
  }

  const external = _recipientViolation(policy, args);
  if (external) {
    const reason =
      `Blocked by communication-agent policy: files share names external `
      + `recipient(s) ${external}`;
    await recordSideEffect('files.share', 'blocked', args.fileId || args.filePath || null, {
      reason, external
    });
    return reason;
  }
  return null;
}

// Teams is a second delivery network with its own membership. A chat or channel
// can contain guests, and the destination's membership is not visible from the
// call arguments -- so this gates the capability rather than pretending to check
// recipients it cannot see.
async function enforceTeamsPolicy(args = {}) {
  const policy = getExecutionPolicy();
  if (!policy) return null;

  const teams = policy.teams || {};
  if (!teams.write) {
    const reason =
      'Blocked by communication-agent policy: Teams writes are not allowed for this '
      + 'agent (a chat or channel may include external guests whose membership this '
      + 'fence cannot see)';
    await recordSideEffect('teams.write', 'blocked', args.chatId || args.channelId || null, { reason });
    return reason;
  }
  return null;
}

module.exports = {
  LEGAL_TOOL_MANIFEST,
  getExecutionPolicy,
  isLegalDraftOnly,
  recordSideEffect,
  enforceOfficeToolPolicy,
  recipientRestrictionActive,
  partitionRecipients,
  enforceMailPolicy,
  enforcePlannerPolicy,
  enforceCalendarPolicy,
  enforceFilesPolicy,
  enforceTeamsPolicy
};
