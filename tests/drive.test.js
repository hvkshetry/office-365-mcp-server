const { describe, it, expect } = require('@jest/globals');
const { filesTools } = require('../files');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

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
