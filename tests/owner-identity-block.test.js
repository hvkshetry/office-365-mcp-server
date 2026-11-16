const { afterEach, describe, expect, it } = require('@jest/globals');

const OWNER = 'hersh@circleh2o.com';
const AGENT = 'admin.agent@circleh2o.com';

function loadPolicy({ blocked, owners } = {}) {
  jest.resetModules();
  delete process.env.LEGAL_AGENT_DRAFT_ONLY;
  delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  if (blocked === undefined) {
    delete process.env.OFFICE_MCP_BLOCK_OWNER_IDENTITY;
  } else {
    process.env.OFFICE_MCP_BLOCK_OWNER_IDENTITY = blocked;
  }
  if (owners === undefined) {
    delete process.env.OFFICE_MCP_OWNER_MAILBOXES;
  } else {
    process.env.OFFICE_MCP_OWNER_MAILBOXES = owners;
  }
  return require('../policy');
}

afterEach(() => {
  delete process.env.OFFICE_MCP_BLOCK_OWNER_IDENTITY;
  delete process.env.OFFICE_MCP_OWNER_MAILBOXES;
  jest.resetModules();
});

const MAIL_WRITES = ['send', 'reply', 'draft', 'draft_reply_all', 'update_draft', 'send_draft'];
const TEAMS_WRITES = [
  ['teams_channel', { operation: 'create_message', teamId: 't', channelId: 'c', content: 'x' }],
  ['teams_channel', { operation: 'reply_to_message', teamId: 't', channelId: 'c', messageId: 'm', content: 'x' }],
  ['teams_chat', { operation: 'send_message', chatId: 'c', content: 'x' }],
  ['teams_chat', { operation: 'update_message', chatId: 'c', messageId: 'm', content: 'x' }],
  ['teams_chat', { operation: 'delete_message', chatId: 'c', messageId: 'm' }],
  ['teams_chat', { operation: 'send', chatId: 'c', content: 'x' }],
  ['teams_chat', { chatId: 'c', content: 'x' }]
];

describe('owner-identity block switched on (OCI deployment)', () => {
  it.each(MAIL_WRITES)('refuses mail %s from the owner mailbox by name', async (operation) => {
    const { enforceOfficeToolPolicy } = loadPolicy({ blocked: '1' });
    for (const mailbox of [undefined, '', 'me', OWNER, ' Hersh@CircleH2O.com ', 'not-an-address',
      '326d3b19-345b-440e-946d-f94e6fbf57f6', 'a@b@circleh2o.com']) {
      const refusal = await enforceOfficeToolPolicy('mail', { operation, mailbox, to: [OWNER] });
      expect(refusal).toContain('OWNER_IDENTITY_BLOCKED');
      expect(refusal).toContain(`mail ${operation}`);
    }
  });

  it.each(MAIL_WRITES)('allows mail %s from an agent mailbox, to the owner', async (operation) => {
    const { enforceOfficeToolPolicy } = loadPolicy({ blocked: 'true' });
    await expect(enforceOfficeToolPolicy('mail', { operation, mailbox: AGENT, to: [OWNER] })).resolves.toBeNull();
  });

  it('keeps every mail read and every non-message Teams call', async () => {
    const { enforceOfficeToolPolicy } = loadPolicy({ blocked: 'yes' });
    for (const operation of ['list', 'read', 'search', 'get_attachment', 'list_drafts', 'focused']) {
      await expect(enforceOfficeToolPolicy('mail', { operation })).resolves.toBeNull();
      await expect(enforceOfficeToolPolicy('mail', { operation, mailbox: OWNER })).resolves.toBeNull();
    }
    await expect(enforceOfficeToolPolicy('teams_channel', { operation: 'list_messages' })).resolves.toBeNull();
    await expect(enforceOfficeToolPolicy('teams_chat', { operation: 'get', chatId: 'c', messageId: 'm' }))
      .resolves.toBeNull();
    await expect(enforceOfficeToolPolicy('teams_chat', { chatId: 'c', messageId: 'm' })).resolves.toBeNull();
    await expect(enforceOfficeToolPolicy('teams_meeting', { operation: 'get_transcript' })).resolves.toBeNull();
  });

  it.each(TEAMS_WRITES)('refuses %s %j as the signed-in owner', async (tool, args) => {
    const { enforceOfficeToolPolicy } = loadPolicy({ blocked: '1' });
    const refusal = await enforceOfficeToolPolicy(tool, args);
    expect(refusal).toContain('OWNER_IDENTITY_BLOCKED');
    expect(refusal).toContain(tool);
  });

  it('reads the owner addresses from OFFICE_MCP_OWNER_MAILBOXES', async () => {
    const { enforceOfficeToolPolicy } = loadPolicy({ blocked: '1', owners: `${OWNER}, alias@circleh2o.com` });
    await expect(enforceOfficeToolPolicy('mail', { operation: 'send', mailbox: 'alias@circleh2o.com' }))
      .resolves.toContain('OWNER_IDENTITY_BLOCKED');
    await expect(enforceOfficeToolPolicy('mail', { operation: 'send', mailbox: AGENT })).resolves.toBeNull();
  });
});

describe('owner-identity block switched off (laptop default)', () => {
  it.each([undefined, '', '0', 'false'])('keeps acting as the owner when the switch is %p', async (blocked) => {
    const { enforceOfficeToolPolicy, isOwnerIdentityBlocked } = loadPolicy({ blocked });
    expect(isOwnerIdentityBlocked()).toBe(false);
    for (const operation of MAIL_WRITES) {
      await expect(enforceOfficeToolPolicy('mail', { operation })).resolves.toBeNull();
    }
    for (const [tool, args] of TEAMS_WRITES) {
      await expect(enforceOfficeToolPolicy(tool, args)).resolves.toBeNull();
    }
  });
});
