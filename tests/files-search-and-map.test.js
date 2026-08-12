/**
 * Regressions for the two confirmed `files` defects.
 *
 * #38 — search issued exactly one /search/query and reported that page's length
 *       as the answer. A Minster manuals folder holding 51 PDFs came back as 20
 *       with nothing indicating the other 31 existed. A manual resolver built on
 *       that silently misses files.
 *
 * #36 — files_map_sharepoint_path built a path from regex heuristics and
 *       announced "mapped successfully" without ever touching the filesystem.
 *       On a host with no sync client every returned path was fiction.
 */

const { describe, it, expect, beforeEach } = require('@jest/globals');

jest.mock('../auth', () => ({ ensureAuthenticated: jest.fn() }));
jest.mock('../utils/graph-api');

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { filesTools } = require('../files');
const config = require('../config');

const filesHandler = filesTools.find(t => t.name === 'files').handler;
const mapHandler = filesTools.find(t => t.name === 'files_map_sharepoint_path').handler;

/** One /search/query page of `n` driveItem hits. */
function page(n, { total, more, startAt = 0 }) {
  return {
    value: [{
      hitsContainers: [{
        total,
        moreResultsAvailable: more,
        hits: Array.from({ length: n }, (_, i) => ({
          resource: {
            id: `id${startAt + i}`,
            name: `manual-${startAt + i}.pdf`,
            webUrl: 'https://example.sharepoint.com/x.pdf',
            lastModifiedDateTime: '2026-08-01T00:00:00Z',
            parentReference: { driveId: 'drive1', path: '/Manuals' }
          }
        }))
      }]
    }]
  };
}

describe('files search pagination (#38)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('mock-access-token');
  });

  it('pages until maxResults is satisfied rather than returning one page', async () => {
    callGraphAPI
      .mockResolvedValueOnce(page(25, { total: 51, more: true, startAt: 0 }))
      .mockResolvedValueOnce(page(25, { total: 51, more: true, startAt: 25 }))
      .mockResolvedValueOnce(page(1, { total: 51, more: false, startAt: 50 }));

    const result = await filesHandler({
      operation: 'search', query: 'manual', fileTypes: ['pdf'], maxResults: 60
    });
    const text = result.content[0].text;

    expect(callGraphAPI).toHaveBeenCalledTimes(3);
    expect(text).toContain('Found 51 files');
    expect(text).toContain('manual-50.pdf');
    expect(text).not.toContain('TRUNCATED');
  });

  it('advances the from offset instead of refetching page one', async () => {
    callGraphAPI
      .mockResolvedValueOnce(page(25, { total: 51, more: true, startAt: 0 }))
      .mockResolvedValueOnce(page(25, { total: 51, more: true, startAt: 25 }))
      .mockResolvedValueOnce(page(1, { total: 51, more: false, startAt: 50 }));

    await filesHandler({ operation: 'search', query: 'manual', maxResults: 60 });

    const froms = callGraphAPI.mock.calls.map(c => c[3].requests[0].from);
    expect(froms).toEqual([0, 25, 50]);
  });

  it('says so when more matches exist than were shown', async () => {
    callGraphAPI.mockResolvedValueOnce(page(20, { total: 51, more: true, startAt: 0 }));

    const result = await filesHandler({
      operation: 'search', query: 'manual', maxResults: 20
    });
    const text = result.content[0].text;

    expect(text).toContain('Showing 20 of 51');
    expect(text).toContain('TRUNCATED');
  });
});

describe('files_map_sharepoint_path existence checking (#36)', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('mock-access-token');
  });

  it('does not report success for a path that does not exist', async () => {
    const result = await mapHandler({
      webUrl: 'https://example.sharepoint.com/sites/Operations/Shared%20Documents/20025%20Danone%20Minster/Manuals/x.pdf'
    });
    const text = result.content[0].text;

    expect(text).not.toContain('mapped successfully');
    expect(text).toContain('could NOT be mapped');
    expect(text).toContain('[not found]');
    expect(text).toMatch(/Do not read the paths below/);
  });

  it('points the caller at the Graph API instead of a fabricated path', async () => {
    const result = await mapHandler({
      webUrl: 'https://example.sharepoint.com/sites/Operations/Shared%20Documents/Manuals/x.pdf'
    });
    expect(result.content[0].text).toContain('operation: "download"');
  });

  it('reports a verified path when the mapped directory really exists', async () => {
    const os = require('os');
    const fs = require('fs');
    const nodePath = require('path');

    const root = fs.mkdtempSync(nodePath.join(os.tmpdir(), 'sp-map-'));
    // The transform turns site "Operations" into "Operations - Documents".
    const mapped = nodePath.join(root, 'Operations - Documents');
    fs.mkdirSync(mapped, { recursive: true });

    const previous = config.SHAREPOINT_SYMLINK_PATH;
    config.SHAREPOINT_SYMLINK_PATH = root;
    try {
      const result = await mapHandler({
        webUrl: 'https://example.sharepoint.com/sites/Operations/Shared%20Documents/'
      });
      const text = result.content[0].text;
      expect(text).toContain('[verified]');
      expect(text).toContain('mapped to an existing local path');
      expect(text).not.toContain('could NOT be mapped');
    } finally {
      config.SHAREPOINT_SYMLINK_PATH = previous;
      fs.rmSync(root, { recursive: true, force: true });
    }
  });
});
