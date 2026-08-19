/**
 * Handler-level tests for the outbound mail fence.
 *
 * These exist because the standalone suite (test-policy-recipients.js) calls
 * enforceMailPolicy() and partitionRecipients() DIRECTLY. That proved the policy
 * functions were correct while `create_rule` was never wired to a caller at all —
 * a dead gate passed 56 assertions. Everything here goes through the real mail
 * tool handler with the real policy module loaded, and asserts on the Graph calls
 * that would actually leave the process.
 */

const { afterEach, beforeEach, describe, expect, it } = require('@jest/globals');

const AGENT = 'engineering.agent@purposeenergy.com';
const INSIDE = 'kuno@purposeenergy.com';
const OUTSIDE = 'vendor@example.com';

const POLICY = JSON.stringify({
  mail: { draft: true, send: true, reply: true, send_draft: true, create_rule: false },
  authorized_mailbox: AGENT,
  allowed_recipient_domains: ['purposeenergy.com']
});

/** Load the mail handler with auth + Graph mocked but the REAL policy module. */
function loadMail(policyJson) {
  jest.resetModules();
  if (policyJson === undefined) delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  else process.env.COMMUNICATION_AGENT_POLICY_JSON = policyJson;

  const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');
  const callGraphAPI = jest.fn().mockResolvedValue({});
  jest.doMock('../auth', () => ({ ensureAuthenticated }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));

  const { emailTools } = require('../email');
  return { mail: emailTools[0].handler, callGraphAPI };
}

/** Graph calls whose method is POST and whose path matches. */
function posts(callGraphAPI, re) {
  return callGraphAPI.mock.calls.filter(c => c[1] === 'POST' && re.test(c[2]));
}

afterEach(() => {
  jest.dontMock('../auth');
  jest.dontMock('../utils/graph-api');
  jest.resetModules();
  delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
});

// ---------------------------------------------------------------- create_rule
describe('create_rule reaches the fence', () => {
  it('refuses to create a rule that forwards outside, and issues no Graph POST', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    const res = await mail({
      operation: 'create_rule',
      mailbox: AGENT,
      displayName: 'exfil',
      forwardTo: [OUTSIDE]
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(posts(callGraphAPI, /messageRules/i)).toHaveLength(0);
  });

  it('refuses even with no forwardTo, because create_rule defaults to deny', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    const res = await mail({
      operation: 'create_rule', mailbox: AGENT, displayName: 'benign', moveToFolder: 'Archive'
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(posts(callGraphAPI, /messageRules/i)).toHaveLength(0);
  });

  it('still refuses an internal forwardTo while create_rule is false', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    await mail({ operation: 'create_rule', mailbox: AGENT, displayName: 'x', forwardTo: [INSIDE] });
    expect(posts(callGraphAPI, /messageRules/i)).toHaveLength(0);
  });

  it('blocks an external forwardTo even when create_rule is opted in', async () => {
    const opted = JSON.parse(POLICY);
    opted.mail.create_rule = true;
    const { mail, callGraphAPI } = loadMail(JSON.stringify(opted));
    const res = await mail({
      operation: 'create_rule', mailbox: AGENT, displayName: 'exfil', forwardTo: [OUTSIDE]
    });
    expect(res.content[0].text).toMatch(/external recipient/i);
    expect(posts(callGraphAPI, /messageRules/i)).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------- reply
describe('reply never re-expands the thread through Graph', () => {
  const thread = {
    subject: 'Design review',
    sender: { emailAddress: { address: INSIDE } },
    from: { emailAddress: { address: INSIDE } },
    toRecipients: [
      { emailAddress: { address: AGENT } },
      { emailAddress: { address: OUTSIDE } }
    ],
    ccRecipients: [{ emailAddress: { address: 'paul@purposeenergy.com' } }],
    conversationId: 'c1'
  };

  beforeEach(() => { jest.resetModules(); });

  it('replyAll uses /reply with pinned recipients, never /replyAll', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    callGraphAPI.mockImplementation(async (_t, method, path) => {
      if (method === 'GET' && /\/messages\//.test(path)) return thread;
      if (method === 'GET') return { mail: AGENT, userPrincipalName: AGENT };
      return {};
    });

    const res = await mail({
      operation: 'reply', mailbox: AGENT, emailId: 'MSG1', body: 'ack', replyAll: true
    });

    // The bug this guards: filtering locally but still calling replyAll lets
    // Graph rebuild the full recipient list server-side.
    expect(posts(callGraphAPI, /\/replyAll$/)).toHaveLength(0);
    expect(posts(callGraphAPI, /createReplyAll/)).toHaveLength(0);

    const sent = posts(callGraphAPI, /\/reply$/);
    expect(sent).toHaveLength(1);
    const addressed = JSON.stringify(sent[0][3]);
    expect(addressed).toContain(INSIDE);
    expect(addressed).not.toContain(OUTSIDE);
    expect(res.content[0].text).toMatch(/removed/i);
    expect(res.content[0].text).toContain(OUTSIDE);
  });

  it('sends nothing when every recipient on the thread is outside', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    callGraphAPI.mockImplementation(async (_t, method, path) => {
      if (method === 'GET' && /\/messages\//.test(path)) {
        return {
          subject: 'x',
          sender: { emailAddress: { address: OUTSIDE } },
          from: { emailAddress: { address: OUTSIDE } },
          toRecipients: [{ emailAddress: { address: AGENT } }],
          ccRecipients: []
        };
      }
      if (method === 'GET') return { mail: AGENT, userPrincipalName: AGENT };
      return {};
    });

    const res = await mail({
      operation: 'reply', mailbox: AGENT, emailId: 'MSG2', body: 'ack', replyAll: true
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(posts(callGraphAPI, /\/reply/)).toHaveLength(0);
  });

  it('is inert with no policy: replyAll still goes to /replyAll', async () => {
    const { mail, callGraphAPI } = loadMail(undefined);
    await mail({ operation: 'reply', mailbox: AGENT, emailId: 'M', body: 'b', replyAll: true });
    expect(posts(callGraphAPI, /\/replyAll$/)).toHaveLength(1);
  });
});

// ----------------------------------------------------------------- send_draft
describe('send_draft verifies the server-side draft', () => {
  it('refuses a draft addressed outside and never POSTs /send', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    callGraphAPI.mockImplementation(async (_t, method) => {
      if (method === 'GET') {
        return { toRecipients: [{ emailAddress: { address: OUTSIDE } }], ccRecipients: [], bccRecipients: [] };
      }
      return {};
    });

    const res = await mail({ operation: 'send_draft', mailbox: AGENT, draftId: 'D1' });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(posts(callGraphAPI, /\/send$/)).toHaveLength(0);
  });

  it('sends a draft whose recipients are all allowed', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    callGraphAPI.mockImplementation(async (_t, method) => {
      if (method === 'GET') {
        return { toRecipients: [{ emailAddress: { address: INSIDE } }], ccRecipients: [], bccRecipients: [] };
      }
      return {};
    });

    await mail({ operation: 'send_draft', mailbox: AGENT, draftId: 'D2' });
    expect(posts(callGraphAPI, /\/send$/)).toHaveLength(1);
  });
});

// ----------------------------------------------------------------------- send
describe('direct send still fences', () => {
  it('blocks an external recipient and issues no sendMail', async () => {
    const { mail, callGraphAPI } = loadMail(POLICY);
    const res = await mail({
      operation: 'send', mailbox: AGENT, to: [OUTSIDE], subject: 's', body: 'b'
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(posts(callGraphAPI, /sendMail/i)).toHaveLength(0);
  });
});

// ============================================================================
// Non-mail outbound paths. Each of these reaches a person without ever touching
// the mail tool, so the recipient fence above does not see them.
// ============================================================================

function loadTool(modPath, exportName, policyJson) {
  jest.resetModules();
  if (policyJson === undefined) delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  else process.env.COMMUNICATION_AGENT_POLICY_JSON = policyJson;

  const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');
  const callGraphAPI = jest.fn().mockResolvedValue({});
  jest.doMock('../auth', () => ({ ensureAuthenticated }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));

  const mod = require(modPath);
  const tools = exportName ? mod[exportName] : null;
  return { handler: tools ? tools[0].handler : mod, callGraphAPI };
}

describe('calendar invitations are fenced', () => {
  const CAL = JSON.stringify({
    mail: { send: true }, authorized_mailbox: AGENT,
    allowed_recipient_domains: ['purposeenergy.com'],
    calendar: { write: true }
  });

  it('blocks an external attendee — Graph would mail them an invitation', async () => {
    const { handler, callGraphAPI } = loadTool('../calendar', 'calendarTools', CAL);
    const res = await handler({
      operation: 'create', subject: 'Design review',
      start: '2026-09-01T15:00:00', end: '2026-09-01T16:00:00',
      attendees: [INSIDE, OUTSIDE]
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('allows an all-internal attendee list', async () => {
    const { handler, callGraphAPI } = loadTool('../calendar', 'calendarTools', CAL);
    await handler({
      operation: 'create', subject: 'Internal sync',
      start: '2026-09-01T15:00:00', end: '2026-09-01T16:00:00', attendees: [INSIDE]
    });
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST').length).toBeGreaterThan(0);
  });

  it('denies calendar writes when the policy grants no calendar capability', async () => {
    const { handler, callGraphAPI } = loadTool('../calendar', 'calendarTools', POLICY);
    const res = await handler({
      operation: 'create', subject: 'x',
      start: '2026-09-01T15:00:00', end: '2026-09-01T16:00:00', attendees: [INSIDE]
    });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('never gates reads', async () => {
    const { handler } = loadTool('../calendar', 'calendarTools', POLICY);
    const res = await handler({ operation: 'list' });
    expect(res.content[0].text).not.toMatch(/Blocked by communication-agent policy/);
  });
});

describe('file sharing is fenced', () => {
  it('blocks share when the policy grants no files capability', async () => {
    const { handler, callGraphAPI } = loadTool('../files', 'filesTools', POLICY);
    const res = await handler({ operation: 'share', fileId: 'F1', recipients: [INSIDE] });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('blocks an anonymous link even when sharing is enabled', async () => {
    const shareOk = JSON.stringify({
      mail: { send: true }, authorized_mailbox: AGENT,
      allowed_recipient_domains: ['purposeenergy.com'],
      files: { share: true, allow_anonymous_links: false }
    });
    const { handler, callGraphAPI } = loadTool('../files', 'filesTools', shareOk);
    const res = await handler({ operation: 'share', fileId: 'F1' });   // no recipients
    expect(res.content[0].text).toMatch(/anyone holding it/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('blocks an external share recipient', async () => {
    const shareOk = JSON.stringify({
      mail: { send: true }, authorized_mailbox: AGENT,
      allowed_recipient_domains: ['purposeenergy.com'],
      files: { share: true, allow_anonymous_links: false }
    });
    const { handler, callGraphAPI } = loadTool('../files', 'filesTools', shareOk);
    const res = await handler({ operation: 'share', fileId: 'F1', recipients: [OUTSIDE] });
    expect(res.content[0].text).toMatch(/external recipient/i);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });
});

describe('Teams writes are fenced', () => {
  it('blocks send_message when the policy grants no teams capability', async () => {
    const { handler, callGraphAPI } = loadTool('../teams/consolidated/teams_chat', null, POLICY);
    const res = await handler({ operation: 'send_message', chatId: '19:abc', content: 'hi' });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('blocks add_member', async () => {
    const { handler, callGraphAPI } = loadTool('../teams/consolidated/teams_chat', null, POLICY);
    const res = await handler({ operation: 'add_member', chatId: '19:abc', userId: 'u1' });
    expect(res.content[0].text).toMatch(/Blocked by communication-agent policy/);
    expect(callGraphAPI.mock.calls.filter(c => c[1] === 'POST')).toHaveLength(0);
  });

  it('never gates reads', async () => {
    const { handler } = loadTool('../teams/consolidated/teams_chat', null, POLICY);
    const res = await handler({ operation: 'list_messages', chatId: '19:abc' });
    expect(JSON.stringify(res)).not.toMatch(/Blocked by communication-agent policy/);
  });
});
