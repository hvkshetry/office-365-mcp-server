const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Consolidated Planner Module - Reduces 18 tools to 8 tools
 * Handles all Planner operations with unified handlers
 */

// 1. Plan operations handler
async function handlePlanOperations(args) {
  const { operation, planId, title, groupId } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        const plans = await callGraphAPI(accessToken, 'GET', 'me/planner/plans', null);
        
        if (!plans.value || plans.value.length === 0) {
          return { content: [{ type: "text", text: "No Planner plans found." }] };
        }
        
        const plansList = plans.value.map(plan => `- ${plan.title} (ID: ${plan.id})`).join('\n');
        return {
          content: [{ type: "text", text: `Found ${plans.value.length} plans:\n\n${plansList}` }]
        };
      
      case 'get':
        if (!planId) throw new Error("Plan ID required for get operation");
        
        const plan = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}`, null);
        
        const planInfo = [
          `Title: ${plan.title}`,
          `ID: ${plan.id}`,
          `Created: ${new Date(plan.createdDateTime).toLocaleString()}`,
          `Owner: ${plan.owner}`
        ].join('\n');
        
        return { content: [{ type: "text", text: `Plan Details:\n\n${planInfo}` }] };
      
      case 'create':
        if (!title || !groupId) throw new Error("Title and groupId required for create operation");
        
        const newPlan = await callGraphAPI(accessToken, 'POST', 'planner/plans', {
          owner: groupId,
          title: title
        });
        
        return {
          content: [{ type: "text", text: `Plan created successfully!\nPlan ID: ${newPlan.id}` }]
        };
      
      case 'update':
        if (!planId) throw new Error("Plan ID required for update operation");
        
        const currentPlan = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}`, null);
        
        const updateData = {};
        if (title) updateData.title = title;
        
        await callGraphAPI(accessToken, 'PATCH', `planner/plans/${planId}`, updateData, {
          'If-Match': currentPlan['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Plan updated successfully!" }] };
      
      case 'delete':
        if (!planId) throw new Error("Plan ID required for delete operation");
        
        const planToDelete = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}`, null);
        
        await callGraphAPI(accessToken, 'DELETE', `planner/plans/${planId}`, null, {
          'If-Match': planToDelete['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Plan deleted successfully!" }] };
      
      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  } catch (error) {
    console.error('Error in plan operations:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 2. Task operations handler
async function handleTaskOperations(args) {
  const { operation, planId, taskId, title, bucketId, dueDateTime, assignedTo, percentComplete } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        if (!planId) throw new Error("Plan ID required for list operation");
        
        const tasks = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}/tasks`, null);
        
        if (!tasks.value || tasks.value.length === 0) {
          return { content: [{ type: "text", text: "No tasks found in this plan." }] };
        }
        
        const tasksList = tasks.value.map(task => {
          const status = task.percentComplete === 100 ? '✓' : task.percentComplete + '%';
          return `- [${status}] ${task.title} (ID: ${task.id})`;
        }).join('\n');
        
        return {
          content: [{ type: "text", text: `Found ${tasks.value.length} tasks:\n\n${tasksList}` }]
        };
      
      case 'create':
        if (!planId || !title) throw new Error("Plan ID and title required for create operation");
        
        const taskData = { planId, title };
        
        if (bucketId) taskData.bucketId = bucketId;
        if (dueDateTime) taskData.dueDateTime = dueDateTime;
        if (percentComplete !== undefined) taskData.percentComplete = percentComplete;
        
        if (assignedTo) {
          taskData.assignments = {};
          const assignees = Array.isArray(assignedTo) ? assignedTo : [assignedTo];
          
          for (const userId of assignees) {
            taskData.assignments[userId] = { orderHint: ' !' };
          }
        }
        
        const newTask = await callGraphAPI(accessToken, 'POST', 'planner/tasks', taskData);
        
        return {
          content: [{ type: "text", text: `Task created successfully!\nTask ID: ${newTask.id}` }]
        };
      
      case 'update':
        if (!taskId) throw new Error("Task ID required for update operation");
        
        const currentTask = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        const updateData = {};
        if (title) updateData.title = title;
        if (percentComplete !== undefined) updateData.percentComplete = percentComplete;
        if (dueDateTime) updateData.dueDateTime = dueDateTime;
        if (bucketId) updateData.bucketId = bucketId;
        
        await callGraphAPI(accessToken, 'PATCH', `planner/tasks/${taskId}`, updateData, null, {
          'If-Match': currentTask['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Task updated successfully!" }] };
      
      case 'delete':
        if (!taskId) throw new Error("Task ID required for delete operation");
        
        const taskToDelete = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        await callGraphAPI(accessToken, 'DELETE', `planner/tasks/${taskId}`, null, null, {
          'If-Match': taskToDelete['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Task deleted successfully!" }] };
      
      case 'get':
        if (!taskId) throw new Error("Task ID required for get operation");
        
        const task = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        const taskInfo = [
          `Title: ${task.title}`,
          `ID: ${task.id}`,
          `Status: ${task.percentComplete}% complete`,
          `Created: ${new Date(task.createdDateTime).toLocaleString()}`,
          `Due: ${task.dueDateTime ? new Date(task.dueDateTime).toLocaleString() : 'Not set'}`,
          `Bucket ID: ${task.bucketId || 'Not assigned'}`
        ].join('\n');
        
        return { content: [{ type: "text", text: `Task Details:\n\n${taskInfo}` }] };
      
      case 'assign':
        if (!taskId || !assignedTo) throw new Error("Task ID and user ID required for assign operation");
        
        const taskToAssign = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        const assignmentData = { assignments: { ...taskToAssign.assignments } };
        const userId = Array.isArray(assignedTo) ? assignedTo[0] : assignedTo;
        
        assignmentData.assignments[userId] = { 
          orderHint: ' !',
          '@odata.type': '#microsoft.graph.plannerAssignment' 
        };
        
        await callGraphAPI(accessToken, 'PATCH', `planner/tasks/${taskId}`, assignmentData, null, {
          'If-Match': taskToAssign['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Task assigned successfully!" }] };
      
      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  } catch (error) {
    console.error('Error in task operations:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 3. Bucket operations handler
async function handleBucketOperations(args) {
  const { operation, planId, bucketId, name, orderHint } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        if (!planId) throw new Error("Plan ID required for list operation");
        
        const buckets = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}/buckets`, null);
        
        if (!buckets.value || buckets.value.length === 0) {
          return { content: [{ type: "text", text: "No buckets found in this plan." }] };
        }
        
        const bucketsList = buckets.value.map(bucket => `- ${bucket.name} (ID: ${bucket.id})`).join('\n');
        
        return {
          content: [{ type: "text", text: `Found ${buckets.value.length} buckets:\n\n${bucketsList}` }]
        };
      
      case 'create':
        if (!planId || !name) throw new Error("Plan ID and name required for create operation");
        
        const bucketData = {
          name: name,
          planId: planId,
          orderHint: orderHint || ' !'
        };
        
        const newBucket = await callGraphAPI(accessToken, 'POST', 'planner/buckets', bucketData);
        
        return {
          content: [{ type: "text", text: `Bucket created successfully!\nBucket ID: ${newBucket.id}` }]
        };
      
      case 'update':
        if (!bucketId) throw new Error("Bucket ID required for update operation");
        
        const currentBucket = await callGraphAPI(accessToken, 'GET', `planner/buckets/${bucketId}`, null);
        
        const updateData = {};
        if (name) updateData.name = name;
        
        await callGraphAPI(accessToken, 'PATCH', `planner/buckets/${bucketId}`, updateData, {
          'If-Match': currentBucket['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Bucket updated successfully!" }] };
      
      case 'delete':
        if (!bucketId) throw new Error("Bucket ID required for delete operation");
        
        const bucketToDelete = await callGraphAPI(accessToken, 'GET', `planner/buckets/${bucketId}`, null);
        
        await callGraphAPI(accessToken, 'DELETE', `planner/buckets/${bucketId}`, null, {
          'If-Match': bucketToDelete['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Bucket deleted successfully!" }] };
      
      case 'get_tasks':
        if (!bucketId) throw new Error("Bucket ID required for get_tasks operation");
        
        const bucketTasks = await callGraphAPI(accessToken, 'GET', `planner/buckets/${bucketId}/tasks`, null);
        
        if (!bucketTasks.value || bucketTasks.value.length === 0) {
          return { content: [{ type: "text", text: "No tasks found in this bucket." }] };
        }
        
        const tasksList = bucketTasks.value.map(task => {
          const status = task.percentComplete === 100 ? '✓' : task.percentComplete + '%';
          return `- [${status}] ${task.title} (ID: ${task.id})`;
        }).join('\n');
        
        return {
          content: [{ type: "text", text: `Found ${bucketTasks.value.length} tasks in bucket:\n\n${tasksList}` }]
        };
      
      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  } catch (error) {
    console.error('Error in bucket operations:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 4. User lookup handler
async function handleUserLookup(args) {
  const { email, emails } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    if (email) {
      // Single user lookup
      const user = await callGraphAPI(accessToken, 'GET', `users/${email}`, null, {
        $select: 'id,displayName,mail'
      });
      
      return {
        content: [{ 
          type: "text", 
          text: `User found:\nName: ${user.displayName}\nEmail: ${user.mail}\nID: ${user.id}` 
        }]
      };
    } else if (emails && Array.isArray(emails)) {
      // Multiple users lookup
      const results = [];
      const errors = [];
      
      for (const userEmail of emails) {
        try {
          const user = await callGraphAPI(accessToken, 'GET', `users/${userEmail}`, null, {
            $select: 'id,displayName,mail'
          });
          
          results.push({
            email: userEmail,
            id: user.id,
            displayName: user.displayName
          });
        } catch (error) {
          errors.push({
            email: userEmail,
            error: error.message
          });
        }
      }
      
      let responseText = '';
      
      if (results.length > 0) {
        responseText += 'Users found:\n';
        results.forEach(user => {
          responseText += `- ${user.displayName} (${user.email}): ${user.id}\n`;
        });
      }
      
      if (errors.length > 0) {
        responseText += '\nUsers not found:\n';
        errors.forEach(err => {
          responseText += `- ${err.email}: ${err.error}\n`;
        });
      }
      
      return { content: [{ type: "text", text: responseText }] };
    } else {
      throw new Error("Either email or emails parameter is required");
    }
  } catch (error) {
    console.error('Error in user lookup:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 5. Enhanced task operations handler
async function handleTaskEnhanced(args) {
  const { operation, planId, taskId, title, bucketId, dueDateTime, assignedTo, percentComplete, removeAssignments } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'create':
        if (!planId || !title) throw new Error("Plan ID and title required for create operation");
        
        const taskData = { planId, title };
        
        if (bucketId) taskData.bucketId = bucketId;
        if (dueDateTime) taskData.dueDateTime = dueDateTime;
        if (percentComplete !== undefined) taskData.percentComplete = percentComplete;
        
        if (assignedTo) {
          taskData.assignments = {};
          const assignees = Array.isArray(assignedTo) ? assignedTo : [assignedTo];
          
          for (const userId of assignees) {
            taskData.assignments[userId] = { orderHint: ' !' };
          }
        }
        
        const newTask = await callGraphAPI(accessToken, 'POST', 'planner/tasks', taskData);
        
        return {
          content: [{ 
            type: "text", 
            text: `Task created successfully with ID: ${newTask.id}` 
          }]
        };
      
      case 'update_assignments':
        if (!taskId) throw new Error("Task ID required for update_assignments operation");
        
        const currentTask = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        const updateData = { assignments: { ...currentTask.assignments } };
        
        if (assignedTo) {
          const assignees = Array.isArray(assignedTo) ? assignedTo : [assignedTo];
          for (const userId of assignees) {
            updateData.assignments[userId] = {
              orderHint: ' !',
              '@odata.type': '#microsoft.graph.plannerAssignment'
            };
          }
        }
        
        if (removeAssignments) {
          const toRemove = Array.isArray(removeAssignments) ? removeAssignments : [removeAssignments];
          for (const userId of toRemove) {
            updateData.assignments[userId] = null;
          }
        }
        
        await callGraphAPI(accessToken, 'PATCH', `planner/tasks/${taskId}`, updateData, null, {
          'If-Match': currentTask['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Task assignments updated successfully!" }] };
      
      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  } catch (error) {
    console.error('Error in enhanced task operations:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 6. Task assignments handler
async function handleTaskAssignments(args) {
  const { operation, taskId, userId, userIds } = args;
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'get':
        if (!taskId) throw new Error("Task ID required for get operation");
        
        const task = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        
        const assignments = task.assignments || {};
        const assigneeList = Object.keys(assignments)
          .filter(id => assignments[id] !== null)
          .map(id => ({
            userId: id,
            orderHint: assignments[id].orderHint
          }));
        
        return {
          content: [{ 
            type: "text", 
            text: `Task: ${task.title}\nAssignments: ${assigneeList.length}\n${assigneeList.map(a => `- User ID: ${a.userId}`).join('\n')}` 
          }]
        };
      
      case 'update':
        if (!taskId) throw new Error("Task ID required for update operation");
        
        const currentTask = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
        const updateData = { assignments: { ...currentTask.assignments } };
        
        if (userId || userIds) {
          const assignees = userIds ? userIds : [userId];
          for (const id of assignees) {
            updateData.assignments[id] = {
              orderHint: ' !',
              '@odata.type': '#microsoft.graph.plannerAssignment'
            };
          }
        }
        
        await callGraphAPI(accessToken, 'PATCH', `planner/tasks/${taskId}`, updateData, null, {
          'If-Match': currentTask['@odata.etag']
        });
        
        return { content: [{ type: "text", text: "Task assignments updated successfully!" }] };
      
      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  } catch (error) {
    console.error('Error in task assignments:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 7. Task details handler
async function handleTaskDetails(args) {
  const { taskId } = args;
  
  if (!taskId) {
    return {
      content: [{ type: "text", text: "Missing required parameter: taskId" }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    // Get task details
    const task = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
    
    // Get task details (description, checklist, etc.)
    const taskDetails = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}/details`, null);
    
    const taskInfo = [
      `Title: ${task.title}`,
      `ID: ${task.id}`,
      `Status: ${task.percentComplete}% complete`,
      `Created: ${new Date(task.createdDateTime).toLocaleString()}`,
      `Due: ${task.dueDateTime ? new Date(task.dueDateTime).toLocaleString() : 'Not set'}`,
      `Bucket ID: ${task.bucketId || 'Not assigned'}`,
      `Description: ${taskDetails.description || 'None'}`,
      `Checklist items: ${taskDetails.checklist ? Object.keys(taskDetails.checklist).length : 0}`,
      `References: ${taskDetails.references ? Object.keys(taskDetails.references).length : 0}`
    ].join('\n');
    
    return {
      content: [{ type: "text", text: `Task Details:\n\n${taskInfo}` }]
    };
  } catch (error) {
    console.error('Error getting task details:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// 8. Bulk operations handler
async function handleBulkOperations(args) {
  const { operation, taskIds, updates } = args;
  
  if (!taskIds || !Array.isArray(taskIds)) {
    return {
      content: [{ type: "text", text: "Missing required parameter: taskIds (array)" }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    const results = { success: [], failed: [] };
    
    for (const taskId of taskIds) {
      try {
        switch (operation) {
          case 'update':
            const currentTask = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
            
            const updateData = {};
            if (updates.percentComplete !== undefined) updateData.percentComplete = updates.percentComplete;
            if (updates.bucketId) updateData.bucketId = updates.bucketId;
            if (updates.dueDateTime) updateData.dueDateTime = updates.dueDateTime;
            
            await callGraphAPI(accessToken, 'PATCH', `planner/tasks/${taskId}`, updateData, null, {
              'If-Match': currentTask['@odata.etag']
            });
            
            results.success.push(taskId);
            break;
          
          case 'delete':
            const taskToDelete = await callGraphAPI(accessToken, 'GET', `planner/tasks/${taskId}`, null);
            
            await callGraphAPI(accessToken, 'DELETE', `planner/tasks/${taskId}`, null, null, {
              'If-Match': taskToDelete['@odata.etag']
            });
            
            results.success.push(taskId);
            break;
          
          default:
            throw new Error(`Unknown operation: ${operation}`);
        }
      } catch (error) {
        results.failed.push({ taskId, error: error.message });
      }
    }
    
    const resultText = [
      `Operation: ${operation}`,
      `Successful: ${results.success.length}`,
      `Failed: ${results.failed.length}`,
      '',
      results.success.length > 0 ? `Success IDs: ${results.success.join(', ')}` : '',
      results.failed.length > 0 ? `Failed:\n${results.failed.map(f => `- ${f.taskId}: ${f.error}`).join('\n')}` : ''
    ].filter(line => line).join('\n');
    
    return {
      content: [{ type: "text", text: resultText }]
    };
  } catch (error) {
    console.error('Error in bulk operations:', error);
    return { content: [{ type: "text", text: `Error: ${error.message}` }] };
  }
}

// Export consolidated tools
const plannerTools = [
  {
    name: 'planner_plan',
    description: 'Handle plan operations (list, get, create, update, delete)',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['list', 'get', 'create', 'update', 'delete'],
          description: 'Operation to perform' 
        },
        planId: { type: 'string', description: 'Plan ID (required for get, update, delete)' },
        title: { type: 'string', description: 'Plan title (required for create, optional for update)' },
        groupId: { type: 'string', description: 'Microsoft 365 group ID (required for create)' }
      },
      required: ['operation']
    },
    handler: handlePlanOperations
  },
  {
    name: 'planner_task',
    description: 'Handle task operations (list, create, update, delete, get, assign)',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['list', 'create', 'update', 'delete', 'get', 'assign'],
          description: 'Operation to perform' 
        },
        planId: { type: 'string', description: 'Plan ID (required for list, create)' },
        taskId: { type: 'string', description: 'Task ID (required for update, delete, get, assign)' },
        title: { type: 'string', description: 'Task title' },
        bucketId: { type: 'string', description: 'Bucket ID' },
        dueDateTime: { type: 'string', description: 'Due date in ISO 8601 format' },
        assignedTo: { 
          oneOf: [
            { type: 'string' },
            { type: 'array', items: { type: 'string' } }
          ],
          description: 'User ID(s) to assign' 
        },
        percentComplete: { type: 'number', description: 'Completion percentage (0-100)' }
      },
      required: ['operation']
    },
    handler: handleTaskOperations
  },
  {
    name: 'planner_bucket',
    description: 'Handle bucket operations (list, create, update, delete, get_tasks)',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['list', 'create', 'update', 'delete', 'get_tasks'],
          description: 'Operation to perform' 
        },
        planId: { type: 'string', description: 'Plan ID (required for list, create)' },
        bucketId: { type: 'string', description: 'Bucket ID (required for update, delete, get_tasks)' },
        name: { type: 'string', description: 'Bucket name' },
        orderHint: { type: 'string', description: 'Order hint for bucket positioning' }
      },
      required: ['operation']
    },
    handler: handleBucketOperations
  },
  {
    name: 'planner_user',
    description: 'Handle user lookup operations (get single or multiple user IDs)',
    inputSchema: {
      type: 'object',
      properties: {
        email: { type: 'string', description: 'Single user email address' },
        emails: { 
          type: 'array', 
          items: { type: 'string' },
          description: 'Array of user email addresses' 
        }
      }
    },
    handler: handleUserLookup
  },
  {
    name: 'planner_task_enhanced',
    description: 'Enhanced task operations with better assignment handling',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['create', 'update_assignments'],
          description: 'Operation to perform' 
        },
        planId: { type: 'string', description: 'Plan ID (required for create)' },
        taskId: { type: 'string', description: 'Task ID (required for update_assignments)' },
        title: { type: 'string', description: 'Task title' },
        bucketId: { type: 'string', description: 'Bucket ID' },
        dueDateTime: { type: 'string', description: 'Due date in ISO 8601 format' },
        assignedTo: { 
          oneOf: [
            { type: 'string' },
            { type: 'array', items: { type: 'string' } }
          ],
          description: 'User ID(s) to assign' 
        },
        removeAssignments: { 
          oneOf: [
            { type: 'string' },
            { type: 'array', items: { type: 'string' } }
          ],
          description: 'User ID(s) to unassign' 
        },
        percentComplete: { type: 'number', description: 'Completion percentage (0-100)' }
      },
      required: ['operation']
    },
    handler: handleTaskEnhanced
  },
  {
    name: 'planner_assignments',
    description: 'Handle task assignments (get, update)',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['get', 'update'],
          description: 'Operation to perform' 
        },
        taskId: { type: 'string', description: 'Task ID' },
        userId: { type: 'string', description: 'Single user ID to assign' },
        userIds: { 
          type: 'array', 
          items: { type: 'string' },
          description: 'Multiple user IDs to assign' 
        }
      },
      required: ['operation', 'taskId']
    },
    handler: handleTaskAssignments
  },
  {
    name: 'planner_task_details',
    description: 'Get detailed task information',
    inputSchema: {
      type: 'object',
      properties: {
        taskId: { type: 'string', description: 'Task ID' }
      },
      required: ['taskId']
    },
    handler: handleTaskDetails
  },
  {
    name: 'planner_bulk_operations',
    description: 'Handle bulk operations for tasks',
    inputSchema: {
      type: 'object',
      properties: {
        operation: { 
          type: 'string', 
          enum: ['update', 'delete'],
          description: 'Operation to perform' 
        },
        taskIds: { 
          type: 'array', 
          items: { type: 'string' },
          description: 'Array of task IDs to process' 
        },
        updates: {
          type: 'object',
          properties: {
            percentComplete: { type: 'number', description: 'Completion percentage (0-100)' },
            bucketId: { type: 'string', description: 'Bucket ID' },
            dueDateTime: { type: 'string', description: 'Due date in ISO 8601 format' }
          },
          description: 'Update data (for update operation)'
        }
      },
      required: ['operation', 'taskIds']
    },
    handler: handleBulkOperations
  }
];

module.exports = { plannerTools };