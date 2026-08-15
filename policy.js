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
  getExecutionPolicy,
  isLegalDraftOnly,
  recordSideEffect,
  enforceOfficeToolPolicy,
  enforceMailPolicy,
  enforcePlannerPolicy,
  enforceCalendarPolicy
};
