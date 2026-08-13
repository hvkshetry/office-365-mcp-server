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
    await handleEmailSearch({ from: 'nikita@circlecbg.com' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets.length).toBeGreaterThan(0);
    expect(gets[0].queryParams.$search).toContain('from:nikita@circlecbg.com');
    expect(gets[0].queryParams.$filter).toBeUndefined();
    expect(anyFilterIncludes(gets, 'from/emailAddress/address')).toBe(false);
  });

  it('routes a to-only search through $search, never a toRecipients $filter', async () => {
    const { handleEmailSearch, callGraphAPI } = loadSearchWithMocks();
    await handleEmailSearch({ to: 'hersh@circleh2o.com' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('to:hersh@circleh2o.com');
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
    await handleEmailSearch({ from: 'nikita@circlecbg.com', startDate: '2026-06-01', endDate: '2026-07-31' });

    const gets = messageGetCalls(callGraphAPI);
    expect(gets[0].queryParams.$search).toContain('from:nikita@circlecbg.com');
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
