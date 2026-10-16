// Exercise the actual tools/list and tools/call registrations with the installed
// SDK client, which validates structured results against advertised schemas.
const { beforeEach, afterEach, describe, expect, it } = require('@jest/globals');

let client;
let graph;
let signals;

beforeEach(async () => {
  jest.resetModules();
  jest.spyOn(console, 'error').mockImplementation(() => {});
  signals = Object.fromEntries(['SIGTERM', 'SIGINT'].map(signal => [signal, process.listeners(signal)]));
  graph = jest.fn();
  jest.doMock('../auth', () => ({
    authTools: [],
    ensureAuthenticated: jest.fn().mockResolvedValue('fixture-token')
  }));
  jest.doMock('../utils/graph-api', () => ({ callGraphAPI: graph }));
  jest.doMock('../utils/audit', () => ({ auditLog: jest.fn() }));

  const { InMemoryTransport } = require('@modelcontextprotocol/sdk/inMemory.js');
  const { Client } = require('@modelcontextprotocol/sdk/client/index.js');
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  jest.doMock('@modelcontextprotocol/sdk/server/stdio.js', () => ({
    StdioServerTransport: jest.fn(() => serverTransport)
  }));
  require('../config').TRANSPORT_TYPE = 'stdio';
  require('../index');
  client = new Client({ name: 'mail-output-test', version: '1.0.0' });
  await client.connect(clientTransport);
});

afterEach(async () => {
  await client?.close();
  for (const [signal, prior] of Object.entries(signals)) {
    for (const listener of process.listeners(signal)) {
      if (!prior.includes(listener)) process.removeListener(signal, listener);
    }
  }
  jest.restoreAllMocks();
});

describe('mail structured output on the MCP wire', () => {
  it('advertises a schema and preserves search IDs through SDK validation', async () => {
    const { tools } = await client.listTools();
    expect(tools.find(tool => tool.name === 'mail').outputSchema.required).toContain('text');
    graph.mockResolvedValue({ value: [{
      id: 'wire-message-id', subject: 'Fixture', from: 'sender@example.com',
      receivedDateTime: '2026-10-01T00:00:00Z', hasAttachments: false
    }] });
    const result = await client.callTool({ name: 'mail', arguments: {
      operation: 'search', from: 'sender@example.com'
    } });
    expect(result.structuredContent.emails[0].id).toBe('wire-message-id');
    expect(result.structuredContent.text).toBe(result.content[0].text);
  });

  it.each(['list', 'search', 'list_drafts'])('also validates empty %s responses', async operation => {
    await client.listTools();
    graph.mockResolvedValue({ value: [] });
    const result = await client.callTool({ name: 'mail', arguments: {
      operation, from: 'sender@example.com'
    } });
    expect(result.structuredContent.text).toBe(result.content[0].text);
    expect(result.structuredContent.text).toMatch(/No (emails|drafts)/);
  });

  it('preserves an authentication error in both representations', async () => {
    await client.listTools();
    require('../auth').ensureAuthenticated.mockRejectedValue(new Error('fixture auth failure'));
    const result = await client.callTool({ name: 'mail', arguments: { operation: 'search' } });
    expect(result.isError).toBe(true);
    expect(result.structuredContent.text).toContain('fixture auth failure');
  });
});
