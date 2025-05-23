const { describe, it, expect, jest } = require('@jest/globals');
const {
  handleListFiles,
  handleGetFileContent,
  handleUploadFile,
  handleCreateFolder,
  handleRestoreRecycleBinItem
} = require('../drive');
const tokenManager = require('../auth/token-manager');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth/token-manager');
jest.mock('../utils/graph-api');

describe('Drive Module', () => {
  const mockTokens = {
    access_token: 'mock-access-token',
    email: 'user@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    tokenManager.loadTokenCache.mockReturnValue(mockTokens);
  });

  describe('handleListFiles', () => {
    it('should list files in root directory', async () => {
      const mockFiles = {
        value: [
          { id: 'file1', name: 'Document.docx', size: 1024 },
          { id: 'folder1', name: 'Projects', folder: {} }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockFiles);
      
      const result = await handleListFiles({});
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/me/drive/root/children',
        null
      );
      expect(result.content[0].text).toContain('Found 2 items');
      expect(result.content[0].text).toContain('Document.docx [File]');
      expect(result.content[0].text).toContain('Projects [Folder]');
    });

    it('should list files in specific folder', async () => {
      const mockFiles = { value: [] };
      
      callGraphAPI.mockResolvedValue(mockFiles);
      
      const result = await handleListFiles({ folderId: 'folder123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/me/drive/items/folder123/children',
        null
      );
      expect(result.content[0].text).toBe('No files found in this location.');
    });
  });

  describe('handleGetFileContent', () => {
    it('should retrieve file content', async () => {
      const fileContent = 'Hello, World!';
      
      callGraphAPI.mockResolvedValue(fileContent);
      
      const result = await handleGetFileContent({ fileId: 'file123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/me/drive/items/file123/content',
        null
      );
      expect(result.content[0].text).toBe('Hello, World!');
    });
  });

  describe('handleUploadFile', () => {
    it('should upload a file successfully', async () => {
      const mockResponse = {
        id: 'newfile123',
        name: 'uploaded.txt',
        size: 256
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleUploadFile({
        fileName: 'uploaded.txt',
        content: 'File content',
        parentFolderId: 'folder123'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'PUT',
        '/me/drive/items/folder123:/uploaded.txt:/content',
        'File content',
        { 'Content-Type': 'text/plain' }
      );
      expect(result.content[0].text).toContain('Successfully uploaded');
    });

    it('should validate required parameters', async () => {
      const result = await handleUploadFile({ fileName: 'test.txt' });
      
      expect(result.content[0].text).toContain('Missing required parameters');
    });
  });

  describe('handleCreateFolder', () => {
    it('should create a folder successfully', async () => {
      const mockResponse = {
        id: 'newfolder123',
        name: 'New Folder'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleCreateFolder({
        folderName: 'New Folder',
        parentFolderId: 'root'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/me/drive/items/root/children',
        {
          name: 'New Folder',
          folder: {}
        }
      );
      expect(result.content[0].text).toContain('Successfully created folder');
    });
  });

  describe('handleRestoreRecycleBinItem', () => {
    it('should restore an item from recycle bin', async () => {
      const mockResponse = {
        id: 'restored123',
        name: 'Restored File.docx'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleRestoreRecycleBinItem({ itemId: 'deleted123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/me/drive/items/deleted123/restore',
        {}
      );
      expect(result.content[0].text).toContain('Successfully restored item');
    });
  });
});