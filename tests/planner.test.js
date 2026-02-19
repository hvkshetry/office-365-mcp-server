const { plannerTools } = require('../planner');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

jest.mock('../auth', () => ({
  ensureAuthenticated: jest.fn()
}));
jest.mock('../utils/graph-api');

// The consolidated planner tool
const plannerHandler = plannerTools[0].handler;

describe('Planner Module (Consolidated)', () => {
  const mockAccessToken = 'mock-access-token';

  beforeEach(() => {
    jest.clearAllMocks();
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
  });

  describe('plan entity', () => {
    it('should list plans', async () => {
      const mockPlans = {
        value: [
          { id: 'plan1', title: 'Project Alpha' },
          { id: 'plan2', title: 'Project Beta' }
        ]
      };

      callGraphAPI.mockResolvedValue(mockPlans);

      const result = await plannerHandler({ entity: 'plan', operation: 'list' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'GET', 'me/planner/plans', null
      );
      expect(result.content[0].text).toContain('Found 2 plans');
    });

    it('should create a plan', async () => {
      callGraphAPI.mockResolvedValue({ id: 'newplan123', title: 'New Project' });

      const result = await plannerHandler({
        entity: 'plan', operation: 'create', title: 'New Project', groupId: 'group456'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'POST', 'planner/plans', { owner: 'group456', title: 'New Project' }
      );
      expect(result.content[0].text).toContain('Plan created successfully');
    });

    it('should update plan with ETag in headers', async () => {
      callGraphAPI.mockResolvedValueOnce({ '@odata.etag': 'W/"planEtag"' });
      callGraphAPI.mockResolvedValueOnce({});

      const result = await plannerHandler({
        entity: 'plan', operation: 'update', planId: 'plan1', title: 'Updated'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'PATCH', 'planner/plans/plan1',
        { title: 'Updated' }, null, { 'If-Match': 'W/"planEtag"' }
      );
      expect(result.content[0].text).toContain('Plan updated successfully');
    });

    it('should delete plan with ETag in headers', async () => {
      callGraphAPI.mockResolvedValueOnce({ '@odata.etag': 'W/"planEtag2"' });
      callGraphAPI.mockResolvedValueOnce({});

      const result = await plannerHandler({
        entity: 'plan', operation: 'delete', planId: 'plan1'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'DELETE', 'planner/plans/plan1',
        null, null, { 'If-Match': 'W/"planEtag2"' }
      );
      expect(result.content[0].text).toContain('Plan deleted successfully');
    });
  });

  describe('task entity', () => {
    it('should list tasks', async () => {
      callGraphAPI.mockResolvedValue({
        value: [
          { id: 'task1', title: 'Design UI', percentComplete: 50 },
          { id: 'task2', title: 'Implement API', percentComplete: 0 }
        ]
      });

      const result = await plannerHandler({ entity: 'task', operation: 'list', planId: 'plan123' });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'GET', 'planner/plans/plan123/tasks', null
      );
      expect(result.content[0].text).toContain('Found 2 tasks');
    });

    it('should create a task', async () => {
      callGraphAPI.mockResolvedValue({ id: 'newtask123' });

      const result = await plannerHandler({
        entity: 'task', operation: 'create', planId: 'plan123', title: 'New Task', bucketId: 'b1'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'POST', 'planner/tasks',
        { planId: 'plan123', title: 'New Task', bucketId: 'b1' }
      );
      expect(result.content[0].text).toContain('Task created successfully');
    });

    it('should update task with ETag in headers (not query params)', async () => {
      callGraphAPI.mockResolvedValueOnce({ '@odata.etag': 'W/"etag123"' });
      callGraphAPI.mockResolvedValueOnce({});

      const result = await plannerHandler({
        entity: 'task', operation: 'update', taskId: 'task123', percentComplete: 100
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'PATCH', 'planner/tasks/task123',
        { percentComplete: 100 }, null, { 'If-Match': 'W/"etag123"' }
      );
      expect(result.content[0].text).toContain('Task updated successfully');
    });

    it('should get task details via get_details operation', async () => {
      callGraphAPI.mockResolvedValueOnce({
        title: 'Test Task', id: 't1', percentComplete: 50,
        createdDateTime: '2025-01-01T00:00:00Z', bucketId: 'b1'
      });
      callGraphAPI.mockResolvedValueOnce({
        description: 'A test task', checklist: {}, references: {}
      });

      const result = await plannerHandler({
        entity: 'task', operation: 'get_details', taskId: 't1'
      });

      expect(result.content[0].text).toContain('Task Details');
      expect(result.content[0].text).toContain('Test Task');
    });
  });

  describe('bucket entity', () => {
    it('should list buckets', async () => {
      callGraphAPI.mockResolvedValue({
        value: [{ id: 'b1', name: 'To Do' }, { id: 'b2', name: 'Done' }]
      });

      const result = await plannerHandler({ entity: 'bucket', operation: 'list', planId: 'p1' });
      expect(result.content[0].text).toContain('Found 2 buckets');
    });

    it('should update bucket with ETag in headers', async () => {
      callGraphAPI.mockResolvedValueOnce({ '@odata.etag': 'W/"bEtag"' });
      callGraphAPI.mockResolvedValueOnce({});

      const result = await plannerHandler({
        entity: 'bucket', operation: 'update', bucketId: 'b1', name: 'Updated'
      });

      expect(callGraphAPI).toHaveBeenCalledWith(
        mockAccessToken, 'PATCH', 'planner/buckets/b1',
        { name: 'Updated' }, null, { 'If-Match': 'W/"bEtag"' }
      );
    });
  });

  describe('user entity', () => {
    it('should lookup single user', async () => {
      callGraphAPI.mockResolvedValue({
        displayName: 'Test User', mail: 'test@example.com', id: 'uid1'
      });

      const result = await plannerHandler({
        entity: 'user', operation: 'lookup', email: 'test@example.com'
      });

      expect(result.content[0].text).toContain('Test User');
    });
  });

  describe('routing', () => {
    it('should require entity and operation', async () => {
      const result = await plannerHandler({ entity: 'plan' });
      expect(result.content[0].text).toContain('Missing required parameters');
    });

    it('should reject invalid entity', async () => {
      const result = await plannerHandler({ entity: 'invalid', operation: 'list' });
      expect(result.content[0].text).toContain('Invalid entity');
    });

    it('should reject invalid task operation', async () => {
      const result = await plannerHandler({ entity: 'task', operation: 'invalid_op' });
      expect(result.content[0].text).toContain('Invalid task operation');
    });
  });
});
