const { describe, expect, it } = require('@jest/globals');

// Loads email/search.js with auth + graph-api mocked. callGraphAPI resolves to 0 results
// by default so the handler returns without needing real Graph data.
function loadSearchWithMocks() {
  jest.resetModules();
  const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');
  const callGraphAPI = jest.fn().mockResolvedValue({ value: [] });
  jest.doMock('../auth', () => ({ ensureAuthenticated }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  const { handleEmailSearch } = require('../email/search');
  return { handleEmailSearch, callGraphAPI };
}

// All GET calls that hit a /messages endpoint, with their queryParams (5th positional arg).
function messageGetCalls(callGraphAPI) {
  return callGraphAPI.mock.calls
    .filter((c) => c[1] === 'GET' && typeof c[2] === 'string' && c[2].includes('messages'))
    .map((c) => ({ endpoint: c[2], queryParams: c[4] || {} }));
}

function anyFilterIncludes(gets, needle) {
  return gets.some((g) => typeof g.queryParams.$filter === 'string' && g.queryParams.$filter.includes(needle));
}

describe('email search predicate routing (Graph "too complex" 400 fix)', () => {
  it('routes a from-only search through $search, never a from/emailAddress $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ from: 'sender@example.com' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets.length).toBeGreaterThan(0);
    expect(gets[0].queryParams.$search).toContain('from:sender@example.com');
    expect(gets[0].queryParams.$filter).toBeUndefined();
    expect(anyFilterIncludes(gets, 'from/emailAddress/address')).toBe(false);
  });

  it('routes a to-only search through $search, never a toRecipients $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ to: 'mailbox@example.com' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('to:mailbox@example.com');
    expect(anyFilterIncludes(gets, 'toRecipients')).toBe(false);
  });

  it('routes hasAttachments through $search, never a hasAttachments $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ hasAttachments: true });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('hasattachment:true');
    expect(anyFilterIncludes(gets, 'hasAttachments eq')).toBe(false);
  });

  it('routes importance through $search, never an importance $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ importance: 'high' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('importance:high');
    expect(anyFilterIncludes(gets, "importance eq")).toBe(false);
  });

  it('routes an array-valued to (mail schema) through $search as OR-joined recipients', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ to: ['a@example.com', 'b@example.com'] });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('to:a@example.com');
    expect(gets[0].queryParams.$search).toContain('to:b@example.com');
    expect(anyFilterIncludes(gets, 'toRecipients')).toBe(false);
  });

  it('from + date range still uses $search (KQL), never a from $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ from: 'sender@example.com', startDate: '2026-06-01', endDate: '2026-07-31' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('from:sender@example.com');
    expect(gets[0].queryParams.$search).toContain('received>=2026-06-01');
    expect(anyFilterIncludes(gets, 'from/emailAddress/address')).toBe(false);
  });

  it('date-only search keeps the $filter path with a valid $orderby', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ startDate: '2026-06-01', endDate: '2026-07-31' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toBeUndefined();
    expect(gets[0].queryParams.$filter).toContain('receivedDateTime ge');
    expect(gets[0].queryParams.$filter).toContain('receivedDateTime le');
    expect(gets[0].queryParams.$orderby).toBe('receivedDateTime desc');
  });

  it('isRead-only search keeps the $filter path (not $search)', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ isRead: false });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toBeUndefined();
    expect(gets[0].queryParams.$filter).toContain('isRead eq false');
  });
});

describe('multi-word query tokenisation in the $filter fallback', () => {
  it('ANDs one contains() per term instead of matching the whole string literally', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ query: 'Air Products PRISM membrane' });

    const gets = messageGetCalls(callGraphAPI);
    const filters = gets
      .map((g) => g.queryParams.$filter)
      .filter((f) => typeof f === 'string' && f.includes('contains(subject'));
    expect(filters.length).toBeGreaterThan(0);

    const f = filters[0];
    // Every term is its own predicate...
    for (const term of ['Air', 'Products', 'PRISM', 'membrane']) {
      expect(f).toContain(`contains(subject, '${term}')`);
    }
    // ...ANDed together, and never the whole phrase as one literal substring.
    expect(f).toContain(' and ');
    expect(f).not.toContain("contains(subject, 'Air Products PRISM membrane')");
  });

  it('leaves a single-word query as one predicate', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ query: 'napier' });

    const gets = messageGetCalls(callGraphAPI);
    const f = gets.map((g) => g.queryParams.$filter).find((x) => typeof x === 'string' && x.includes('contains(subject'));
    expect(f).toContain("contains(subject, 'napier')");
  });

  it("escapes single quotes so a term like O'Brien cannot break the filter", async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ query: "O'Brien" });

    const gets = messageGetCalls(callGraphAPI);
    const f = gets.map((g) => g.queryParams.$filter).find((x) => typeof x === 'string' && x.includes('contains(subject'));
    expect(f).toContain("contains(subject, 'O''Brien')");
  });

  it('tokenises the subject parameter the same way', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ subject: 'Draft Tech Proposal' });

    const gets = messageGetCalls(callGraphAPI);
    const f = gets.map((g) => g.queryParams.$filter).find((x) => typeof x === 'string' && x.includes('contains(subject'));
    for (const term of ['Draft', 'Tech', 'Proposal']) {
      expect(f).toContain(`contains(subject, '${term}')`);
    }
  });

  it('says the fallback searched subject lines only when it finds nothing', async () => {
    const { handleEmailSearch } = loadSearchWithMocks();
    const res = await handleEmailSearch({ query: 'somethingthatdoesnotexist' });
    const text = res.content.map((c) => c.text).join('\n');
    // A bare "No emails found" reads as an authoritative empty mailbox; it isn't.
    expect(text).toMatch(/SUBJECT LINES ONLY/i);
    expect(text).toMatch(/bodies were not searched/i);
  });
});

// A subject-only result that LOOKS authoritative is the failure mode: a caller reads it as
// "the mailbox has nothing" when bodies were never read. These assert the degradation is
// always visible, and that a transient $search error is distinguishable from a real empty.
describe('the subject-only degradation announces itself', () => {
  // Graph answers $search with an error, then answers the $filter fallback with hits.
  function loadWithSearchError(message) {
    jest.resetModules();
    const ensureAuthenticated = jest.fn().mockResolvedValue('access-token');
    const callGraphAPI = jest.fn().mockImplementation((_t, _m, _p, _d, params = {}) => {
      if (params.$search) return Promise.reject(new Error(message));
      return Promise.resolve({ value: [{
        id: 'AAA', subject: 'Chiller enquiry', from: { emailAddress: { address: 'sender@example.com' } },
        receivedDateTime: '2026-09-03T07:29:37Z', hasAttachments: false, importance: 'normal', isRead: true
      }] });
    });
    jest.doMock('../auth', () => ({ ensureAuthenticated }));
    jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
    const { handleEmailSearch } = require('../email/search');
    return { handleEmailSearch };
  }

  it('labels a RESULTS list as subject-only and names the transient cause', async () => {
    const { handleEmailSearch } = loadWithSearchError('API call failed with status 401: token expired');
    const res = await handleEmailSearch({ query: 'Chiller', mailbox: 'mailbox@example.com' });
    const text = res.content.map((c) => c.text).join('\n');
    // The old text said only "(using filter fallback)" — easy to read past when hits came back.
    expect(text).toMatch(/SUBJECT LINES ONLY/i);
    expect(text).toMatch(/bodies were NOT searched/i);
    expect(text).toMatch(/401/);
  });

  it('distinguishes "$search matched nothing" from "$search failed"', async () => {
    const { handleEmailSearch } = loadWithSearchError('API call failed with status 429: throttled');
    const failed = (await handleEmailSearch({ query: 'zzz', mailbox: 'mailbox@example.com' }))
      .content.map((c) => c.text).join('\n');
    expect(failed).toMatch(/\$search failed/i);
    expect(failed).not.toMatch(/\$search matched nothing/i);

    // Default harness: $search returns an empty value array, i.e. Graph looked and found nothing.
    const { handleEmailSearch: h2 } = loadSearchWithMocks();
    const empty = (await h2({ query: 'zzz', mailbox: 'mailbox@example.com' }))
      .content.map((c) => c.text).join('\n');
    expect(empty).toMatch(/\$search matched nothing/i);
  });

  it('explains why a from/to search cannot retry on the subject filter', async () => {
    const { handleEmailSearch } = loadWithSearchError('API call failed with status 503');
    const res = await handleEmailSearch({ from: 'sender@example.com', mailbox: 'mailbox@example.com' });
    const text = res.content.map((c) => c.text).join('\n');
    expect(text).toMatch(/503/);
    expect(text).toMatch(/cannot fall back/i);
  });
});
