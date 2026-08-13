/**
 * Tests for the opt-in recipient-domain fence in policy.js.
 *
 * The first assertion group is the important one: with no policy configured —
 * which is every persona in production today — nothing changes. A regression
 * there would silently start blocking admin-assistant's or finance-agent's
 * legitimate outbound mail.
 *
 * Run: node test-policy-recipients.js
 */

const assert = require('assert');
const path = require('path');

const POLICY_PATH = path.join(__dirname, 'policy.js');

function withPolicy(policyObj, fn) {
  // policy.js caches on first read, so the module must be reloaded per policy.
  delete require.cache[require.resolve(POLICY_PATH)];
  if (policyObj === undefined) {
    delete process.env.COMMUNICATION_AGENT_POLICY_JSON;
  } else {
    process.env.COMMUNICATION_AGENT_POLICY_JSON =
      typeof policyObj === 'string' ? policyObj : JSON.stringify(policyObj);
  }
  return fn(require(POLICY_PATH));
}

const INTERNAL = 'tamika@purposeenergy.com';
const VENDOR = 'epentek@membranechemicals.com';
const MBX = 'procurement.agent@purposeenergy.com';

const FENCED = {
  mail: { draft: true, send: true, reply: true, send_draft: false },
  authorized_mailbox: 'procurement.agent@purposeenergy.com',
  allowed_recipient_domains: ['purposeenergy.com']
};

let pass = 0, fail = 0;
async function check(name, got, wantBlocked) {
  const blocked = got !== null && got !== undefined;
  const ok = blocked === wantBlocked;
  ok ? pass++ : fail++;
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) console.log(`        expected ${wantBlocked ? 'BLOCKED' : 'allowed'}, got: ${got}`);
}

(async () => {
  // ---- 1. no policy at all: the state of every persona in production -------
  console.log('\nno policy configured (all existing personas):');
  await withPolicy(undefined, async (p) => {
    await check('send to a vendor is untouched',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), false);
    await check('send_draft is untouched',
      await p.enforceMailPolicy('send_draft', { draftId: 'X' }), false);
    await check('update_draft is untouched',
      await p.enforceMailPolicy('update_draft', { draftId: 'X', to: [VENDOR] }), false);
  });

  // ---- 2. a policy WITHOUT the new key: operation gates only ---------------
  console.log('\npolicy without allowed_recipient_domains (back-compat):');
  await withPolicy({ mail: { send: true, reply: true, draft: true, send_draft: true } }, async (p) => {
    await check('external recipient is NOT blocked (key absent)',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), false);
  });

  // ---- 3. the fenced procurement policy ------------------------------------
  console.log('\nprocurement policy (allowed_recipient_domains set):');
  await withPolicy(FENCED, async (p) => {
    await check('send to internal', await p.enforceMailPolicy('send', { mailbox: MBX, to: [INTERNAL] }), false);
    await check('send to vendor', await p.enforceMailPolicy('send', { mailbox: MBX, to: [VENDOR] }), true);
    await check('DRAFT to vendor is fenced too',
      await p.enforceMailPolicy('draft', { mailbox: MBX, to: [VENDOR] }), true);
    await check('UPDATE_DRAFT swapping in a vendor',
      await p.enforceMailPolicy('update_draft', { mailbox: MBX, draftId: 'X', to: [VENDOR] }), true);
    await check('vendor hidden in cc',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [INTERNAL], cc: [VENDOR] }), true);
    await check('vendor hidden in bcc',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [INTERNAL], bcc: ['x@gmail.com'] }), true);
    await check('send_draft still blocked by the operation gate',
      await p.enforceMailPolicy('send_draft', { mailbox: MBX, draftId: 'X' }), true);

    // shape handling — each of these was a fail-open in an earlier design
    await check('bare-string to, internal',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: INTERNAL }), false);
    await check('bare-string to, vendor',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: VENDOR }), true);
    await check('comma-separated smuggling a vendor',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: `${INTERNAL}, ${VENDOR}` }), true);
    await check('semicolon-separated',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: `${INTERNAL}; ${VENDOR}` }), true);
    await check('display-name form, internal',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [`Tamika Fraser <${INTERNAL}>`] }), false);
    await check('display-name form hiding a vendor',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [`Tamika Fraser <${VENDOR}>`] }), true);
    await check('Graph-style dict, vendor',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [{ emailAddress: { address: VENDOR } }] }), true);
    await check('Graph-style dict, internal',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [{ emailAddress: { address: INTERNAL } }] }), false);
    await check('toRecipients alias',
      await p.enforceMailPolicy('send', { mailbox: MBX, toRecipients: [VENDOR] }), true);

    // domain-parsing traps
    await check('suffix trap purposeenergy.com.evil.co',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['x@purposeenergy.com.evil.co'] }), true);
    await check('subdomain is not the domain',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['x@mail.purposeenergy.com'] }), true);
    await check('case-insensitive internal',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['Tamika@PURPOSEENERGY.COM'] }), false);
    await check('trailing-dot domain',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['x@purposeenergy.com.'] }), false);
    await check('malformed, no @',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['notanemail'] }), true);
    await check('double @',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: ['a@b@purposeenergy.com'] }), true);
    await check('unknown object shape denies',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [{ nope: 1 }] }), true);

    // reads and unaddressed calls
    await check('list is unaffected', await p.enforceMailPolicy('list', { mailbox: MBX }), false);
    await check('reply with no explicit recipient is not blocked HERE',
      await p.enforceMailPolicy('reply', { mailbox: MBX, emailId: 'X' }), false);
  });

  // ---- 4. malformed policy JSON -> fails OPEN (pre-existing behaviour) -----
  console.log('\nmalformed policy JSON:');
  await withPolicy('{not valid json', async (p) => {
    await check('fails OPEN — a typo in the env var disables the fence',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), false);
  });

  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
