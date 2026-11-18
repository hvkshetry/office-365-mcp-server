/**
 * The request recipient gate (owner rulings 2026-11-18, HOUSE-CBG-ORIGINATION-PROMOTER-THREAD-REPLY-20261118 and
 * HOUSE-LEGAL-PROMOTER-GROUP-REPLY-20261118) and the request's own message id, through the path index.js runs:
 * enforceOfficeToolPolicy, then the mail handler, with the real policy module and a fake Graph.
 */
const { afterEach, describe, expect, it } = require('@jest/globals');

const LEGAL = 'legal.agent@circleh2o.com';
const CBG = 'origination.agent@example.org';
const ENGINEERING = 'process.agent@example.org';
const OWNER = 'owner@example.org';
const PROMOTER_A = 'promoter.one@example.org';
const PROMOTER_B = 'promoter.two@example.org';
const OUTSIDER = 'counsel@lawfirm.example';
const ALLOWED = [OWNER, PROMOTER_A, PROMOTER_B, LEGAL, CBG, ENGINEERING];
// A Graph message id is ~152 characters; the live failure dropped 32 of them.
const ORIGINAL_ID = `AAMkAGI2${'x'.repeat(136)}AAA=`;

function address(a) {
  return { emailAddress: { address: a } };
}

function requestPolicy(persona, mailbox, { gate = true, replyEmailId = ORIGINAL_ID } = {}) {
  const ops = ['draft', 'draft_reply_all', 'update_draft', 'send', 'reply', 'send_draft'];
  return {
    reply_actor: persona,
    authorized_mailbox: mailbox,
    mail: Object.fromEntries(ops.map(op => [op, true])),
    planner: { write: false },
    calendar: { write: false },
    ...(replyEmailId ? { reply_email_id: replyEmailId } : {}),
    ...(gate ? { recipient_gate: { allowed_recipients: ALLOWED, ruling_id: 'HOUSE-TEST-GATE' } } : {})
  };
}

function load({ policy, legal = false } = {}) {
  jest.resetModules();
  delete process.env.OFFICE_MCP_BLOCK_OWNER_IDENTITY;
  if (legal) process.env.LEGAL_AGENT_DRAFT_ONLY = '1'; else delete process.env.LEGAL_AGENT_DRAFT_ONLY;
  if (policy === undefined) delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  else process.env.COMMUNICATION_AGENT_POLICY_JSON = JSON.stringify(policy);

  const original = { value: null };
  const callGraphAPI = jest.fn(async (token, method, endpoint, body, query) => {
    if (method === 'GET' && /\/messages\/draft-/.test(endpoint)) return original.draft;
    if (method === 'GET' && endpoint.endsWith(`/messages/${ORIGINAL_ID}`)) return original.value;
    if (method === 'GET' && /\/messages\//.test(endpoint)) throw new Error('API call failed with status 400: Id is malformed.');
    if (endpoint.endsWith('/createReplyAll')) return { id: 'draft-reply-1', toRecipients: [], ccRecipients: [] };
    return {};
  });
  jest.doMock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('access-token') }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  const policyModule = require('../policy');
  const { emailTools } = require('../email');
  // index.js: the tool policy (which may write the request's id into args), then the handler on the same args.
  async function call(args) {
    const refusal = await policyModule.enforceOfficeToolPolicy('mail', args);
    if (refusal) return { refused: refusal, text: refusal };
    const result = await emailTools[0].handler(args);
    return { refused: null, text: result.content[0].text, result };
  }
  return { call, callGraphAPI, original, policy: policyModule };
}

function thread({ from = OWNER, to = [LEGAL, PROMOTER_A], cc = [PROMOTER_B] } = {}) {
  return { subject: 'Question', from: address(from), sender: address(from), replyTo: [],
    toRecipients: to.map(address), ccRecipients: cc.map(address),
    internetMessageId: '<m@example.org>', conversationId: 'conv-1' };
}

const sent = callGraphAPI => callGraphAPI.mock.calls.filter(([, method, endpoint]) =>
  method === 'POST' && /\/(replyAll|reply|send|sendMail)$/.test(endpoint));

afterEach(() => {
  delete process.env.LEGAL_AGENT_DRAFT_ONLY;
  delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  jest.dontMock('../auth');
  jest.dontMock('../utils/graph-api');
  jest.resetModules();
});

describe.each([
  ['legal', LEGAL, true],
  ['cbg-origination', CBG, false]
])('recipient gate: %s', (persona, mailbox, legal) => {
  it('replies all on a promoter-only thread (agent mailboxes count as internal)', async () => {
    const { call, callGraphAPI, original } = load({ policy: requestPolicy(persona, mailbox), legal });
    original.value = thread({ cc: [PROMOTER_B, ENGINEERING] });
    const out = await call({ operation: 'reply', emailId: 'original', replyAll: true, body: '<p>a</p>', mailbox });
    expect(out.refused).toBeNull();
    expect(out.text).toContain('Reply sent successfully');
    expect(sent(callGraphAPI).map(c => c[2])).toEqual([`users/${mailbox}/messages/${ORIGINAL_ID}/replyAll`]);
  });

  it('refuses reply-all when the thread carries an external cc, and the reply-all draft goes ahead', async () => {
    const { call, callGraphAPI, original } = load({ policy: requestPolicy(persona, mailbox), legal });
    original.value = thread({ cc: [PROMOTER_B, OUTSIDER] });
    const out = await call({ operation: 'reply', emailId: ORIGINAL_ID, replyAll: true, body: '<p>a</p>', mailbox });
    expect(out.refused || out.text).toContain('RECIPIENT_GATE');
    expect(out.text).toContain(OUTSIDER);
    expect(out.text).toContain('draft_reply_all');
    expect(sent(callGraphAPI)).toEqual([]);

    const draft = await call({ operation: 'draft_reply_all', emailId: 'original', body: '<p>a</p>', mailbox });
    expect(draft.refused).toBeNull();
    expect(draft.result.draftId).toBe('draft-reply-1');
    expect(sent(callGraphAPI)).toEqual([]);
  });

  it('refuses an external To or Reply-To on the original, an added cc/bcc, a send and a send_draft outside', async () => {
    const { call, callGraphAPI, original } = load({ policy: requestPolicy(persona, mailbox), legal });
    original.value = { ...thread(), replyTo: [address(OUTSIDER)] };
    expect((await call({ operation: 'reply', emailId: 'original', body: 'x', mailbox })).text)
      .toContain('RECIPIENT_GATE');
    original.value = thread();
    expect((await call({ operation: 'reply', emailId: 'original', replyAll: true, cc: [OUTSIDER], body: 'x', mailbox }))
      .text).toContain('RECIPIENT_GATE');
    expect((await call({ operation: 'send', to: [PROMOTER_A], bcc: [OUTSIDER], subject: 's', body: 'x', mailbox })).text)
      .toContain('RECIPIENT_GATE');
    expect((await call({ operation: 'send', to: [`Promoter <${PROMOTER_A}>`], cc: [PROMOTER_B], subject: 's', body: 'x', mailbox }))
      .text).toContain('Email sent successfully');
    original.draft = { toRecipients: [address(PROMOTER_A)], ccRecipients: [], bccRecipients: [address(OUTSIDER)] };
    expect((await call({ operation: 'send_draft', draftId: 'draft-9', mailbox })).text).toContain('RECIPIENT_GATE');
    original.draft = { toRecipients: [address(PROMOTER_A)], ccRecipients: [address(OWNER)], bccRecipients: [] };
    expect((await call({ operation: 'send_draft', draftId: 'draft-9', mailbox })).text).toContain('Draft sent');
    expect(sent(callGraphAPI).map(c => c[2])).toEqual([`users/${mailbox}/sendMail`,
      `users/${mailbox}/messages/draft-9/send`]);
  });

  it('refuses when the recipients cannot be read (fail closed)', async () => {
    const { call, callGraphAPI } = load({ policy: requestPolicy(persona, mailbox, { replyEmailId: null }), legal });
    const out = await call({ operation: 'reply', emailId: 'AAMk-unknown', replyAll: true, body: 'x', mailbox });
    expect(out.text).toContain('RECIPIENT_GATE');
    expect(out.text).toContain('could not be read');
    expect(sent(callGraphAPI)).toEqual([]);
  });

  it('refuses a forwarding rule outside the group', async () => {
    const { call } = load({ policy: requestPolicy(persona, mailbox), legal });
    const out = await call({ operation: 'create_rule', displayName: 'fwd', forwardTo: [OUTSIDER], mailbox });
    expect(out.refused).toContain('RECIPIENT_GATE');
  });
});

describe('Legal without a request gate stays draft-only', () => {
  it('refuses reply, send and send_draft with no policy and with an ungated policy', async () => {
    for (const policy of [undefined, requestPolicy('legal', LEGAL, { gate: false })]) {
      const { call, callGraphAPI, original } = load({ policy, legal: true });
      original.value = thread();
      for (const args of [{ operation: 'reply', emailId: ORIGINAL_ID, replyAll: true, body: 'x' },
        { operation: 'send', to: [PROMOTER_A], subject: 's', body: 'x' }, { operation: 'send_draft', draftId: 'draft-1' }]) {
        const out = await call({ ...args, mailbox: LEGAL });
        expect(out.refused).toMatch(/Legal OS tool manifest/);
      }
      expect(sent(callGraphAPI)).toEqual([]);
    }
  });

  it('refuses a gated Legal transmission from any mailbox but legal.agent@', async () => {
    const { call } = load({ policy: requestPolicy('legal', LEGAL), legal: true });
    const out = await call({ operation: 'reply', emailId: 'original', replyAll: true, body: 'x', mailbox: PROMOTER_A });
    expect(out.refused).toContain(`must explicitly use mailbox "${LEGAL}"`);
  });
});

describe('other personas are unchanged', () => {
  it('a request policy without a gate replies to an external thread as before', async () => {
    const { call, callGraphAPI, original } = load({ policy: requestPolicy('pe-process', ENGINEERING, { gate: false }) });
    original.value = thread({ to: [ENGINEERING], cc: [OUTSIDER] });
    const out = await call({ operation: 'reply', emailId: 'original', replyAll: true, body: 'x', mailbox: ENGINEERING });
    expect(out.text).toContain('Reply sent successfully');
    // no recipient read: the ungated reply makes exactly the one Graph call it always made
    expect(callGraphAPI.mock.calls.map(c => `${c[1]} ${c[2]}`))
      .toEqual([`POST users/${ENGINEERING}/messages/${ORIGINAL_ID}/replyAll`]);
  });
});

describe('the request own message id', () => {
  it('"original" and an omitted emailId on reply / draft_reply_all take the request id', async () => {
    const { policy } = load({ policy: requestPolicy('pe-process', ENGINEERING, { gate: false }) });
    for (const args of [{ operation: 'reply', emailId: 'original' }, { operation: 'reply', emailId: ' Original ' },
      { operation: 'draft_reply_all' }, { operation: 'read', emailId: 'original' },
      { operation: 'get_attachment', emailId: 'original', attachmentId: 'a' }]) {
      expect(policy.applyRequestEmailId('mail', args)).toBeNull();
      expect(args.emailId).toBe(ORIGINAL_ID);
    }
  });

  it('refuses a retyped id with characters missing, naming where the id comes from, before any Graph call', async () => {
    const { call, callGraphAPI } = load({ policy: requestPolicy('pe-process', ENGINEERING, { gate: false }) });
    const retyped = ORIGINAL_ID.slice(0, 60) + ORIGINAL_ID.slice(92);
    for (const operation of ['reply', 'draft_reply_all']) {
      const out = await call({ operation, emailId: retyped, replyAll: true, body: 'x', mailbox: ENGINEERING });
      expect(out.refused).toContain('REQUEST_EMAIL_ID_MISMATCH');
      expect(out.refused).toContain(`(${retyped.length} characters)`);
      expect(out.refused).toContain('emailId: "original"');
      expect(out.refused).toContain('reply_email_id');
    }
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  it('leaves other operations and runs without the key alone; "original" without the key is refused', async () => {
    const { policy } = load({ policy: requestPolicy('pe-process', ENGINEERING, { gate: false }) });
    const read = { operation: 'read', emailId: 'AAMk-other' };
    expect(policy.applyRequestEmailId('mail', read)).toBeNull();
    expect(read.emailId).toBe('AAMk-other');
    expect(policy.applyRequestEmailId('calendar', { operation: 'get', emailId: 'original' })).toBeNull();

    const bare = load({ policy: requestPolicy('pe-process', ENGINEERING, { gate: false, replyEmailId: null }) }).policy;
    const reply = { operation: 'reply', emailId: 'AAMk-any' };
    expect(bare.applyRequestEmailId('mail', reply)).toBeNull();
    expect(reply.emailId).toBe('AAMk-any');
    expect(bare.applyRequestEmailId('mail', { operation: 'reply', emailId: 'original' })).toContain('REQUEST_EMAIL_ID');
    expect(load({ policy: undefined }).policy.applyRequestEmailId('mail', { operation: 'reply', emailId: 'x' }))
      .toBeNull();
  });
});
