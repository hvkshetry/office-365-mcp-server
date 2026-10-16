/**
 * Two ways the files tool answered a question nobody asked.
 *
 * `list` documented `parentPath` but never read it, so naming a folder returned the contents of
 * the drive root instead -- which reads as "that folder is empty" or "it holds the wrong things".
 * A confident wrong answer is worse than an error.
 *
 * And `move` demanded a destination for a pure rename, forcing the caller to look up the folder
 * the item is already in and pass it back unchanged.
 */

const { describe, it, expect, beforeEach } = require('@jest/globals');
const { filesTools } = require('../files');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth', () => ({ ensureAuthenticated: jest.fn() }));
jest.mock('../utils/graph-api');

const filesHandler = filesTools[0].handler;
const DRIVE = 'drive-fixture-1';

describe('list honours parentPath', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    callGraphAPI.mockResolvedValue({ value: [] });
  });

  it('addresses the named folder inside the given drive', async () => {
    await filesHandler({
      operation: 'list',
      parentPath: '02 Project Documents/07 Transmittals',
      driveId: DRIVE
    });
    const endpoint = callGraphAPI.mock.calls[0][2];
    expect(endpoint).toContain(`/drives/${DRIVE}/root:/`);
    expect(endpoint).toContain('07%20Transmittals');
  });

  it('does not silently fall back to the drive root', async () => {
    await filesHandler({ operation: 'list', parentPath: '02 Project Documents', driveId: DRIVE });
    expect(callGraphAPI.mock.calls[0][2]).not.toBe(`/drives/${DRIVE}/root`);
  });

  it('prefers an explicit folderId, which is exact', async () => {
    await filesHandler({
      operation: 'list',
      folderId: '01ITEMID',
      parentPath: 'ignored',
      driveId: DRIVE
    });
    expect(callGraphAPI.mock.calls[0][2]).toContain(`/drives/${DRIVE}/items/01ITEMID`);
  });
});

describe('move renames without a destination', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue('token');
    callGraphAPI.mockResolvedValue({
      name: 'renamed.pdf',
      parentReference: { path: '/drive/root:/02 Project Documents' }
    });
  });

  it('sends only the new name', async () => {
    const result = await filesHandler({
      operation: 'move',
      fileId: '01ITEMID',
      newName: 'renamed.pdf',
      driveId: DRIVE
    });
    const body = callGraphAPI.mock.calls[0][3];
    expect(body.name).toBe('renamed.pdf');
    expect(body.parentReference).toBeUndefined();
    expect(result.content[0].text).toContain('renamed successfully');
  });

  it('still moves when a destination is given', async () => {
    await filesHandler({
      operation: 'move',
      fileId: '01ITEMID',
      destinationId: '01DESTID',
      driveId: DRIVE
    });
    expect(callGraphAPI.mock.calls[0][3].parentReference).toEqual({ id: '01DESTID' });
  });

  it('says so when asked to do neither', async () => {
    const result = await filesHandler({ operation: 'move', fileId: '01ITEMID' });
    expect(result.content[0].text).toContain('Nothing to do');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  it('still requires a fileId', async () => {
    const result = await filesHandler({ operation: 'move', newName: 'x.pdf' });
    expect(result.content[0].text).toContain('Missing required parameter: fileId');
  });
});
