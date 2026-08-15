const { afterEach, describe, expect, it } = require('@jest/globals');

const LEGAL_MAILBOX = 'legal.agent@circleh2o.com';

function loadPolicy({ communicationPolicy } = {}) {
  jest.resetModules();
  process.env.LEGAL_AGENT_DRAFT_ONLY = '1';
  if (communicationPolicy === undefined) {
    delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  } else {
    process.env.COMMUNICATION_AGENT_POLICY_JSON = communicationPolicy;
  }
  return require('../policy');
}

afterEach(() => {
  delete process.env.LEGAL_AGENT_DRAFT_ONLY;
  delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  jest.resetModules();
});

describe('Legal OS Office tool manifest', () => {
  it('allows reviewed reads and blocks unreviewed tools and operations', async () => {
    const { enforceOfficeToolPolicy } = loadPolicy();

    await expect(enforceOfficeToolPolicy('mail', { operation: 'read' })).resolves.toBeNull();
    await expect(enforceOfficeToolPolicy('future_sharepoint_mutator', {
      operation: 'publish'
    })).resolves.toContain('is not reviewed');
    await expect(enforceOfficeToolPolicy('files', {
      operation: 'upload'
    })).resolves.toContain('files.upload is not allowed');
  });

  it('allows only explicit legal-mailbox draft operations', async () => {
    const { enforceOfficeToolPolicy } = loadPolicy();

    await expect(enforceOfficeToolPolicy('mail', {
      operation: 'draft',
      mailbox: LEGAL_MAILBOX
    })).resolves.toBeNull();
    await expect(enforceOfficeToolPolicy('mail', {
      operation: 'draft_reply_all'
    })).resolves.toContain('must explicitly use mailbox');
    await expect(enforceOfficeToolPolicy('mail', {
      operation: 'update_draft',
      mailbox: 'me'
    })).resolves.toContain(LEGAL_MAILBOX);
    await expect(enforceOfficeToolPolicy('mail', {
      operation: 'send',
      mailbox: LEGAL_MAILBOX
    })).resolves.toContain('mail.send is not allowed');
  });

  it('blocks every transmission operation even without a communication policy', async () => {
    const { enforceMailPolicy } = loadPolicy();

    for (const operation of ['send', 'reply', 'send_draft']) {
      await expect(enforceMailPolicy(operation, {
        mailbox: LEGAL_MAILBOX
      })).resolves.toContain('draft-only policy');
    }
  });

  it('fails closed when an injected communication policy is malformed', async () => {
    const { enforceMailPolicy } = loadPolicy({ communicationPolicy: '{not-json' });

    await expect(enforceMailPolicy('draft', {
      mailbox: LEGAL_MAILBOX
    })).resolves.toContain('is invalid');
  });
});
