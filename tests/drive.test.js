const { describe, it, expect } = require('@jest/globals');
const { filesTools } = require('../files');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');
const fs = require('fs');
const path = require('path');

jest.mock('../auth', () => ({
  ensureAuthenticated: jest.fn()
}));
jest.mock('../utils/graph-api');

// The consolidated files tool
const filesHandler = filesTools[0].handler;

describe('Files Module', () => {
  const mockAccessToken = 'mock-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    config.TEMP_FILES_PATH = '/tmp/office-mcp-test-files';
    if (!fs.existsSync(config.TEMP_FILES_PATH)) {
      fs.mkdirSync(config.TEMP_FILES_PATH, { recursive: true });
    }
  });

  describe('routing', () => {
    it('should require operation parameter', async () => {
      const result = await filesHandler({});
      expect(result.content[0].text).toContain('Missing required parameter: operation');
    });
  });

  describe('list operation', () => {
    it('should list files in root directory', async () => {
      const mockFiles = {
        value: [
          { id: 'file1', name: 'Document.docx', size: 1024, file: {} },
          { id: 'folder1', name: 'Projects', folder: { childCount: 3 } }
        ]
      };

      callGraphAPI.mockResolvedValue(mockFiles);

      const result = await filesHandler({ operation: 'list' });

      expect(result.content[0].text).toContain('Document.docx');
      expect(result.content[0].text).toContain('Projects');
    });

    it('should list files in specific folder', async () => {
      callGraphAPI.mockResolvedValue({ value: [] });

      const result = await filesHandler({
        operation: 'list',
        folderId: 'folder123'
      });

      expect(result.content[0].text).toBeTruthy();
    });
  });

  describe('search operation', () => {
    it('should search for files', async () => {
      // Default scope is 'all' which uses /search/query (hitsContainers format)
      const mockResults = {
        value: [{
          hitsContainers: [{
            hits: [
              { resource: { id: 'file1', name: 'budget.xlsx', size: 2048, file: {} } }
            ]
          }]
        }]
      };

      callGraphAPI.mockResolvedValue(mockResults);

      const result = await filesHandler({
        operation: 'search',
        query: 'budget'
      });

      expect(result.content[0].text).toContain('budget.xlsx');
    });
  });

  describe('upload operation', () => {
    it('should upload a file', async () => {
      const mockResponse = {
        id: 'newfile123',
        name: 'uploaded.txt',
        size: 256
      };

      callGraphAPI.mockResolvedValue(mockResponse);

      const result = await filesHandler({
        operation: 'upload',
        fileName: 'uploaded.txt',
        content: 'File content here'
      });

      expect(result.content[0].text).toContain('uploaded');
    });
  });

  describe('binary upload operation', () => {
    it('uses the requested drive root when parentPath is omitted', async () => {
      callGraphAPI.mockResolvedValue({ id: 'file-fixture', name: 'report #1.bin', size: 4 });
      await filesHandler({ operation: 'upload', fileName: 'report #1.bin', contentBase64: 'AAECAw==', driveId: 'drive-fixture' });
      expect(callGraphAPI.mock.calls[0][2]).toBe('/drives/drive-fixture/root:/report%20%231.bin:/content');
    });

    it('rejects malformed base64 before making a request', async () => {
      const result = await filesHandler({ operation: 'upload', fileName: 'fixture.bin', contentBase64: 'AAECAw=!' });
      expect(result.content[0].text).toContain('strict RFC 4648 base64');
      expect(callGraphAPI).not.toHaveBeenCalled();
    });

    it('decodes strict base64 before uploading binary content', async () => {
      callGraphAPI.mockResolvedValue({
        id: 'binary123',
        name: 'drawing.dwg',
        size: 4
      });

      const result = await filesHandler({
        operation: 'upload',
        fileName: 'drawing.dwg',
        contentBase64: 'AAECAw=='
      });

      expect(result.content[0].text).toContain('drawing.dwg');
      expect(callGraphAPI.mock.calls.at(-1)[3]).toEqual(Buffer.from([0, 1, 2, 3]));
    });

    it('rejects simultaneous text and binary upload bodies', async () => {
      const result = await filesHandler({
        operation: 'upload',
        fileName: 'ambiguous.bin',
        content: 'text',
        contentBase64: 'dGV4dA=='
      });

      expect(result.content[0].text).toContain('exactly one');
    });
  });

  describe('download operation', () => {
    it('should return inline content for text-like files', async () => {
      callGraphAPI
        .mockResolvedValueOnce({
          id: 'text123',
          name: 'notes.txt',
          size: 11,
          file: { mimeType: 'text/plain' }
        })
        .mockResolvedValueOnce('hello world');

      const result = await filesHandler({
        operation: 'download',
        fileId: 'text123'
      });

      const text = result.content[0].text;
      expect(text).toContain('File content downloaded successfully');
      expect(text).toContain('hello world');
      expect(text).toContain('Path Scope: local-to-mcp-server-host');
    });

    it('should save binary downloads to local temp files', async () => {
      // Bytes >= 0x80 on purpose. [0,1,2,3] round-trips through
      // Buffer.from(String(buf), 'utf8') unchanged, so a fixture of low bytes
      // passes even when the Buffer branch is bypassed and the lossy
      // String() fallback runs -- the exact corruption this path exists to
      // prevent. High bytes become U+FFFD and the comparison below catches it.
      const BINARY = Buffer.from([0x50, 0x4b, 0x03, 0x04, 0xde, 0xad, 0xbe, 0xef]);
      callGraphAPI
        .mockResolvedValueOnce({
          id: 'binary123',
          name: 'recording.mp4',
          size: BINARY.length,
          file: { mimeType: 'video/mp4' }
        })
        .mockResolvedValueOnce(BINARY);

      const result = await filesHandler({
        operation: 'download',
        fileId: 'binary123',
        driveId: 'drive123'
      });

      const text = result.content[0].text;
      expect(text).toContain('Local Path:');
      expect(text).toContain('MCP Host:');
      expect(text).toContain('Path Scope: local-to-mcp-server-host');
      expect(text).toContain('files { operation: "cleanup_download"');

      const localPath = text.match(/Local Path: (.+)/)[1].trim();
      expect(localPath.startsWith(config.TEMP_FILES_PATH)).toBe(true);
      expect(fs.existsSync(localPath)).toBe(true);
      expect(fs.readFileSync(localPath)).toEqual(BINARY);
      fs.unlinkSync(localPath);
    });

    it('should clean up temp downloads inside the configured temp directory', async () => {
      const tempFile = path.join(config.TEMP_FILES_PATH, 'cleanup-target.bin');
      fs.writeFileSync(tempFile, Buffer.from([1]));

      const result = await filesHandler({
        operation: 'cleanup_download',
        path: tempFile
      });

      expect(result.content[0].text).toContain('Cleaned up downloaded file');
      expect(fs.existsSync(tempFile)).toBe(false);
    });

    it('should reject cleanup outside the configured temp directory', async () => {
      const result = await filesHandler({
        operation: 'cleanup_download',
        path: '/tmp/not-office-mcp-test-files/secret.bin'
      });

      expect(result.content[0].text).toContain('Security: can only clean up files');
    });
  });

  describe('create_folder operation', () => {
    it('should create a folder', async () => {
      const mockResponse = {
        id: 'newfolder123',
        name: 'New Folder',
        folder: {},
        parentReference: { path: '/drive/root:' }
      };

      callGraphAPI.mockResolvedValue(mockResponse);

      const result = await filesHandler({
        operation: 'create_folder',
        name: 'New Folder'
      });

      expect(result.content[0].text).toContain('New Folder');
    });
  });
});
