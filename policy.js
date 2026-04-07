/**
 * Runtime side-effect policy and audit hooks for Office MCP.
 *
 * Policy is injected per communication-agent run through environment variables
 * on the Office MCP process in that agent's ephemeral `.mcp.json`.
 */

let _cachedPolicy;

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
    _cachedPolicy = null;
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
    case 'send':
      return !!mailPolicy.send;
    case 'reply':
      return !!mailPolicy.reply;
    case 'send_draft':
      return !!mailPolicy.send_draft;
    default:
      return true;
  }
}

async function enforceMailPolicy(operation, params = {}) {
  const policy = getExecutionPolicy();
  if (!policy) {
    return null;
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
  getExecutionPolicy,
  recordSideEffect,
  enforceMailPolicy,
  enforcePlannerPolicy,
  enforceCalendarPolicy
};
