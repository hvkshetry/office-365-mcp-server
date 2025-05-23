const { describe, it, expect, jest } = require('@jest/globals');
const {
  handleListPlans,
  handleCreatePlan,
  handleListTasks,
  handleCreateTask,
  handleUpdateTask,
  handleListBuckets,
  handleCreateBucket
} = require('../planner');
const tokenManager = require('../auth/token-manager');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth/token-manager');
jest.mock('../utils/graph-api');

describe('Planner Module', () => {
  const mockTokens = {
    access_token: 'mock-access-token',
    email: 'user@example.com'
  };

  beforeEach(() => {
    jest.clearAllMocks();
    tokenManager.loadTokenCache.mockReturnValue(mockTokens);
  });

  describe('handleListPlans', () => {
    it('should list plans for current user', async () => {
      const mockPlans = {
        value: [
          { id: 'plan1', title: 'Project Alpha' },
          { id: 'plan2', title: 'Project Beta' }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockPlans);
      
      const result = await handleListPlans({});
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/me/planner/plans',
        null
      );
      expect(result.content[0].text).toContain('Found 2 plans');
    });

    it('should list plans for specific group', async () => {
      const mockPlans = { value: [] };
      
      callGraphAPI.mockResolvedValue(mockPlans);
      
      const result = await handleListPlans({ groupId: 'group123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/groups/group123/planner/plans',
        null
      );
      expect(result.content[0].text).toBe('No plans found.');
    });
  });

  describe('handleCreatePlan', () => {
    it('should create a plan successfully', async () => {
      const mockResponse = {
        id: 'newplan123',
        title: 'New Project'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleCreatePlan({
        title: 'New Project',
        ownerId: 'group456'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/planner/plans',
        {
          title: 'New Project',
          owner: 'group456'
        }
      );
      expect(result.content[0].text).toContain('Successfully created plan');
    });

    it('should validate required parameters', async () => {
      const result = await handleCreatePlan({ title: 'Test' });
      
      expect(result.content[0].text).toContain('Missing required parameters');
    });
  });

  describe('handleListTasks', () => {
    it('should list tasks for a plan', async () => {
      const mockTasks = {
        value: [
          { id: 'task1', title: 'Design UI', percentComplete: 50 },
          { id: 'task2', title: 'Implement API', percentComplete: 0 }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockTasks);
      
      const result = await handleListTasks({ planId: 'plan123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/planner/plans/plan123/tasks',
        null
      );
      expect(result.content[0].text).toContain('Found 2 tasks');
      expect(result.content[0].text).toContain('Design UI (50% complete)');
    });
  });

  describe('handleCreateTask', () => {
    it('should create a task successfully', async () => {
      const mockResponse = {
        id: 'newtask123',
        title: 'New Task'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleCreateTask({
        planId: 'plan123',
        title: 'New Task',
        bucketId: 'bucket123'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/planner/tasks',
        {
          planId: 'plan123',
          title: 'New Task',
          bucketId: 'bucket123'
        }
      );
      expect(result.content[0].text).toContain('Successfully created task');
    });
  });

  describe('handleUpdateTask', () => {
    it('should update task completion', async () => {
      // Mocking the GET request for ETag
      callGraphAPI.mockResolvedValueOnce({
        '@odata.etag': 'W/"etag123"'
      });
      
      // Mocking the PATCH request
      callGraphAPI.mockResolvedValueOnce({
        id: 'task123',
        percentComplete: 100
      });
      
      const result = await handleUpdateTask({
        taskId: 'task123',
        percentComplete: 100
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'PATCH',
        '/planner/tasks/task123',
        { percentComplete: 100 },
        { 'If-Match': 'W/"etag123"' }
      );
      expect(result.content[0].text).toContain('Successfully updated task');
    });
  });

  describe('handleListBuckets', () => {
    it('should list buckets for a plan', async () => {
      const mockBuckets = {
        value: [
          { id: 'bucket1', name: 'To Do' },
          { id: 'bucket2', name: 'In Progress' }
        ]
      };
      
      callGraphAPI.mockResolvedValue(mockBuckets);
      
      const result = await handleListBuckets({ planId: 'plan123' });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'GET',
        '/planner/plans/plan123/buckets',
        null
      );
      expect(result.content[0].text).toContain('Found 2 buckets');
    });
  });

  describe('handleCreateBucket', () => {
    it('should create a bucket successfully', async () => {
      const mockResponse = {
        id: 'newbucket123',
        name: 'Done'
      };
      
      callGraphAPI.mockResolvedValue(mockResponse);
      
      const result = await handleCreateBucket({
        planId: 'plan123',
        name: 'Done'
      });
      
      expect(callGraphAPI).toHaveBeenCalledWith(
        'mock-access-token',
        'POST',
        '/planner/buckets',
        {
          planId: 'plan123',
          name: 'Done'
        }
      );
      expect(result.content[0].text).toContain('Successfully created bucket');
    });
  });
});