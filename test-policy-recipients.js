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

  // ---- 3a2. ambiguous address forms are DENIED, not resolved ---------------
  // The permissive parser took the first `<...>` and ignored everything around
  // it, so a string could name one mailbox to this fence and another to
  // Exchange. Ambiguity is now rejected outright.
  console.log('\nambiguous / smuggled address forms:');
  await withPolicy(FENCED, async (p) => {
    const deny = [
      'attacker@evil.example <tamika@purposeenergy.com>',   // display name posing as an address
      'Display <tamika@purposeenergy.com> <attacker@evil.example>', // two bracket pairs
      'Display <tamika@purposeenergy.com> trailing-junk',   // junk after the closing bracket
      '<tamika@purposeenergy.com><x@evil.example>',
      'tamika@purposeenergy.com>extra',
      'tamika@purposeenergy.com <',
      'tamika@purposeenergy',                               // no dot in the domain
      'tamika@.purposeenergy.com',                          // leading dot
      '@purposeenergy.com',                                 // empty local part
      'tamika@',
    ];
    for (const addr of deny) {
      await check(`denied: ${JSON.stringify(addr)}`,
        await p.enforceMailPolicy('send', { mailbox: MBX, to: [addr] }), true);
    }
    const allow = [
      INTERNAL,
      'Tamika Fraser <tamika@purposeenergy.com>',
      'TAMIKA@PurposeEnergy.COM',
      '  tamika@purposeenergy.com  ',
    ];
    for (const addr of allow) {
      await check(`allowed: ${JSON.stringify(addr)}`,
        await p.enforceMailPolicy('send', { mailbox: MBX, to: [addr] }), false);
    }
  });

  // ---- 3b. per-address allowlist (mirrors the inbound lockdown model) ------
  console.log('\nper-address allowlist (finance-agent shape):');
  const FMBX = 'finance.agent@purposeenergy.com';
  const FIN = {
    mail: { draft: true, send: true, reply: true, send_draft: false },
    authorized_mailbox: 'finance.agent@purposeenergy.com',
    allowed_recipient_addresses: [
      'hersh@purposeenergy.com', 'pgerry@purposeenergy.com',
      'svalley@purposeenergy.com', 'fitch@purposeenergy.com'
    ]
  };
  await withPolicy(FIN, async (p) => {
    await check('named person allowed',
      await p.enforceMailPolicy('send', { mailbox: FMBX, to: ['hersh@purposeenergy.com'] }), false);
    await check('case-insensitive named person',
      await p.enforceMailPolicy('send', { mailbox: FMBX, to: ['Hersh@PurposeEnergy.com'] }), false);
    await check('display-name form of a named person',
      await p.enforceMailPolicy('send', { mailbox: FMBX, to: ['Paul Gerry <pgerry@purposeenergy.com>'] }), false);
    await check('SAME DOMAIN but not on the list is blocked',
      await p.enforceMailPolicy('send', { mailbox: FMBX, to: ['someoneelse@purposeenergy.com'] }), true);
    await check('vendor blocked',
      await p.enforceMailPolicy('send', { mailbox: FMBX, to: [VENDOR] }), true);
    await check('allowed person + unlisted colleague in cc is blocked',
      await p.enforceMailPolicy('send',
        { mailbox: FMBX, to: ['hersh@purposeenergy.com'], cc: ['someoneelse@purposeenergy.com'] }), true);
  });

  console.log('\ndomain + address together (operations-assistant shape):');
  const OPS = {
    mail: { draft: true, send: true, reply: true, send_draft: false },
    authorized_mailbox: 'operations.agent@purposeenergy.com',
    allowed_recipient_domains: ['purposeenergy.com'],
    allowed_recipient_addresses: ['jessica.mcgovern@uvm.edu']
  };
  await withPolicy(OPS, async (p) => {
    await check('anyone on the domain',
      await p.enforceMailPolicy('send', { mailbox: 'operations.agent@purposeenergy.com', to: ['kuno@purposeenergy.com'] }), false);
    await check('the named external collaborator',
      await p.enforceMailPolicy('send', { mailbox: 'operations.agent@purposeenergy.com', to: ['jessica.mcgovern@uvm.edu'] }), false);
    await check('a different address at the same external domain is blocked',
      await p.enforceMailPolicy('send', { mailbox: 'operations.agent@purposeenergy.com', to: ['someone.else@uvm.edu'] }), true);
    await check('vendor still blocked',
      await p.enforceMailPolicy('send', { mailbox: 'operations.agent@purposeenergy.com', to: [VENDOR] }), true);
  });

  // ---- 3c. create_rule: an inbox rule can auto-forward past the fence ------
  console.log('\ncreate_rule gate:');
  await withPolicy(FENCED, async (p) => {
    await check('create_rule denied by default under any policy',
      await p.enforceMailPolicy('create_rule', { mailbox: MBX }), true);
    await check('move is not gated (mailbox hygiene, no exfiltration path)',
      await p.enforceMailPolicy('move', { mailbox: MBX }), false);
  });
  await withPolicy({ ...FENCED, mail: { ...FENCED.mail, create_rule: true } }, async (p) => {
    await check('create_rule allowed when explicitly opted in',
      await p.enforceMailPolicy('create_rule', { mailbox: MBX }), false);
  });

  // ---- 3d. partitionRecipients: the reply-all strip ------------------------
  console.log('\npartitionRecipients (reply-all strip):');
  const rcpt = (a) => ({ emailAddress: { address: a, name: a } });
  await withPolicy(FENCED, async (p) => {
    const mixed = p.partitionRecipients([rcpt(INTERNAL), rcpt(VENDOR), rcpt('kuno@purposeenergy.com')]);
    await check('enforced flag set', mixed.enforced ? null : 'not enforced', false);
    await check('keeps both internal recipients', mixed.kept.length === 2 ? null : `kept ${mixed.kept.length}`, false);
    await check('removes exactly the vendor',
      (mixed.removed.length === 1 && mixed.removed[0] === VENDOR) ? null : `removed ${mixed.removed}`, false);

    const allExternal = p.partitionRecipients([rcpt(VENDOR)]);
    await check('all-external thread keeps nobody (caller must refuse)',
      allExternal.kept.length === 0 ? null : 'kept someone', false);

    const junk = p.partitionRecipients([{ emailAddress: {} }, { nope: 1 }]);
    await check('unparseable recipients are removed, never kept',
      junk.kept.length === 0 ? null : 'kept junk', false);
  });
  await withPolicy(undefined, async (p) => {
    const untouched = p.partitionRecipients([rcpt(VENDOR)]);
    await check('no policy: nothing is stripped and enforced is false',
      (!untouched.enforced && untouched.kept.length === 1) ? null : 'altered without a policy', false);
  });

  // ---- 4. malformed policy JSON -> fails OPEN (pre-existing behaviour) -----
  console.log('\nmalformed policy JSON:');
  await withPolicy('{not valid json', async (p) => {
    await check('fails OPEN — a typo in the env var disables the fence',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), false);
  });

  // ---- 5. COMMUNICATION_AGENT_POLICY_REQUIRED: the Phase 3 flip -----------
  // Until this is set, a persona with no policy behaves exactly as it always
  // has. Once set, absent OR malformed refuses to act instead of acting freely.
  console.log('\nfail-closed flip (COMMUNICATION_AGENT_POLICY_REQUIRED=1):');
  process.env.COMMUNICATION_AGENT_POLICY_REQUIRED = '1';
  await withPolicy(undefined, async (p) => {
    await check('send refused when no policy is configured',
      await p.enforceMailPolicy('send', { to: [INTERNAL] }), true);
    await check('reply refused', await p.enforceMailPolicy('reply', { emailId: 'X' }), true);
    await check('create_rule refused', await p.enforceMailPolicy('create_rule', {}), true);
    await check('READS still work — a missing policy stops acting, not looking',
      await p.enforceMailPolicy('list', {}), false);
  });
  await withPolicy('{not valid json', async (p) => {
    await check('malformed policy now refuses instead of failing open',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), true);
  });
  await withPolicy(FENCED, async (p) => {
    await check('a valid policy is unaffected by the flip — internal allowed',
      await p.enforceMailPolicy('send', { mailbox: MBX, to: [INTERNAL] }), false);
    await check('a valid policy is unaffected by the flip — vendor blocked',
      await p.enforceMailPolicy('send', { to: [VENDOR] }), true);
  });
  delete process.env.COMMUNICATION_AGENT_POLICY_REQUIRED;

  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})();
