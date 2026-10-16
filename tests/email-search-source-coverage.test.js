const { describe, expect, it, afterEach } = require('@jest/globals');

const message = {
  id: 'fixture-id', subject: 'Drawing review',
  from: { emailAddress: { address: 'engineer@example.com' } },
  receivedDateTime: '2026-09-21T12:00:00Z', isRead: true, importance: 'normal'
};

function setup() {
  jest.resetModules();
  jest.spyOn(console, 'error').mockImplementation(() => {});
  const callGraphAPI = jest.fn().mockResolvedValue({ value: [message] });
  jest.doMock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('fixture-token') }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI }));
  return { callGraphAPI, search: require('../email/search').handleEmailSearch };
}

afterEach(() => jest.restoreAllMocks());

describe('mail search results survive clients that prefer structuredContent', () => {
  const emails = [
    { ...message, id: 'message-one', hasAttachments: true },
    { ...message, id: 'message-two', subject: 'Second result', hasAttachments: false }
  ];

  it.each(['microsoft-search', 'graph-search', 'subject-filter'])(
    'includes every ID and the listing in %s results', async (method) => {
      const { search, callGraphAPI } = setup();
      let args;
      if (method === 'microsoft-search') {
        callGraphAPI.mockResolvedValue({ value: [{ hitsContainers: [{
          hits: emails.map(resource => ({ resource })), total: 2
        }] }] });
        args = { query: 'drawing', useRelevance: true };
      } else {
        callGraphAPI.mockResolvedValue({ value: emails });
        args = method === 'graph-search'
          ? { from: 'engineer@example.com' }
          : { subject: 'Drawing' };
      }

      const result = await search(args);
      expect(result.structuredContent).toMatchObject({ method, returned: 2 });
      expect(result.structuredContent.text).toBe(result.content[0].text);
      expect(result.structuredContent.emails).toEqual(emails.map(email => ({
        id: email.id,
        subject: email.subject,
        from: email.from.emailAddress.address,
        receivedDateTime: email.receivedDateTime,
        hasAttachments: email.hasAttachments
      })));
      for (const email of emails) {
        expect(result.structuredContent.text).toContain(email.id);
      }
    }
  );

  it('uses the same Microsoft Search hitId fallback in both representations', async () => {
    const { search, callGraphAPI } = setup();
    const { id, ...resource } = message;
    callGraphAPI.mockResolvedValue({ value: [{ hitsContainers: [{
      hits: [{ hitId: 'fallback-message-id', resource }]
    }] }] });
    const result = await search({ query: 'drawing', useRelevance: true });
    expect(result.structuredContent.emails[0].id).toBe('fallback-message-id');
    expect(result.structuredContent.text).toContain('ID: fallback-message-id');
  });
});

describe('source-review searches preserve full-text predicates and disclose bounds', () => {
  it('escapes KQL phrase quotes inside the Graph search string', async () => {
    const { search, callGraphAPI } = setup();
    await search({ query: 'Banthar', mailbox: 'engineer@example.com' });
    const encoded = callGraphAPI.mock.calls[0][4].$search;
    expect(JSON.parse(encoded)).toBe('(subject:"Banthar" OR body:"Banthar" OR from:"Banthar")');
    expect(encoded).toContain('\\"Banthar\\"');
  });

  it('keeps text plus dates on full-text search, including body matches', async () => {
    const { search, callGraphAPI } = setup();
    await search({ query: 'Sterling', startDate: '2026-09-01', mailbox: 'engineer@example.com' });
    const params = callGraphAPI.mock.calls[0][4];
    expect(JSON.parse(params.$search)).toContain('body:"Sterling"');
    expect(JSON.parse(params.$search)).toContain('received>=2026-09-01');
    expect(params.$filter).toBeUndefined();
  });

  it('does not reinterpret failed explicit KQL as subject terms', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockRejectedValue(new Error('service unavailable'));
    const result = await search({ query: 'Banthar AND from:shekhar', mailbox: 'engineer@example.com' });
    expect(result.isError).toBe(true);
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
    expect(result.content[0].text).toContain('not evidence of absence');
  });

  it('reports a genuine empty KQL result without broadening the query', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockResolvedValue({ value: [] });
    const result = await search({ query: 'subject:fixture AND from:nobody', mailbox: 'engineer@example.com' });
    expect(result.isError).toBe(false);
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  it('marks a failed sender-only search as an error, not an empty mailbox', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockRejectedValue(new Error('service unavailable'));
    const result = await search({ from: 'engineer@example.com' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('Email search failed');
    expect(callGraphAPI).toHaveBeenCalledTimes(1);
  });

  it('marks a nextLink result partial even below the requested limit', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockResolvedValue({ value: [message], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' });
    const result = await search({ query: 'Banthar', mailbox: 'engineer@example.com', maxResults: 1000 });
    expect(result.structuredContent).toMatchObject({ limited: true, hasMore: true, pageLimit: 250, returned: 1, serviceLimit: 1000 });
    expect(result.content[0].text).toContain('PARTIAL SEARCH');
  });

  it('marks a filled page partial even when the service omits nextLink', async () => {
    const { search } = setup();
    const result = await search({ query: 'Banthar', mailbox: 'engineer@example.com', maxResults: 1 });
    expect(result.structuredContent.limited).toBe(true);
  });

  it('discloses the remaining simple-query subject-only fallback', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockRejectedValueOnce(new Error('invalid search'));
    const result = await search({ query: 'Banthar', mailbox: 'engineer@example.com' });
    expect(result.content[0].text).toContain('SUBJECT LINES ONLY');
    expect(result.structuredContent.method).toBe('subject-filter');
  });

  it('supports a subject predicate with no query field', async () => {
    const { search, callGraphAPI } = setup();
    const result = await search({ from: 'engineer@example.com', subject: 'Control Philosophy' });
    expect(result.isError).not.toBe(true);
    expect(JSON.parse(callGraphAPI.mock.calls[0][4].$search)).toContain('subject:"Control Philosophy"');
  });

  it('reports Microsoft Search pagination too', async () => {
    const { search, callGraphAPI } = setup();
    callGraphAPI.mockResolvedValue({ value: [{ hitsContainers: [{ hits: [{ resource: message }], moreResultsAvailable: true, total: 8 }] }] });
    const result = await search({ query: 'Banthar', useRelevance: true });
    expect(result.structuredContent).toMatchObject({ method: 'microsoft-search', hasMore: true, limited: true });
  });
});
