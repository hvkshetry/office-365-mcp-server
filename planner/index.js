const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { safeTool } = require('../utils/errors');
const { enforcePlannerPolicy, recordSideEffect } = require('../policy');

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
        
        await callGraphAPI(accessToken, 'PATCH', `planner/plans/${planId}`, updateData, null, {
          'If-Match': currentPlan['@odata.etag']
        });

        return { content: [{ type: "text", text: "Plan updated successfully!" }] };

      case 'delete':
        if (!planId) throw new Error("Plan ID required for delete operation");

        const planToDelete = await callGraphAPI(accessToken, 'GET', `planner/plans/${planId}`, null);

        await callGraphAPI(accessToken, 'DELETE', `planner/plans/${planId}`, null, null, {
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
            taskData.assignments[userId] = { 
              '@odata.type': '#microsoft.graph.plannerAssignment',
              orderHint: ' !' 
            };
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
        
        await callGraphAPI(accessToken, 'PATCH', `planner/buckets/${bucketId}`, updateData, null, {
          'If-Match': currentBucket['@odata.etag']
        });

        return { content: [{ type: "text", text: "Bucket updated successfully!" }] };

      case 'delete':
        if (!bucketId) throw new Error("Bucket ID required for delete operation");

        const bucketToDelete = await callGraphAPI(accessToken, 'GET', `planner/buckets/${bucketId}`, null);

        await callGraphAPI(accessToken, 'DELETE', `planner/buckets/${bucketId}`, null, null, {
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
            taskData.assignments[userId] = { 
              '@odata.type': '#microsoft.graph.plannerAssignment',
              orderHint: ' !' 
            };
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

/**
 * Unified planner handler - single entry point for ALL planner operations
 * Uses entity + operation routing to consolidate 8 tools into 1
 */
async function handlePlanner(args) {
  if (!args || typeof args !== 'object') {
    return {
      content: [{
        type: "text",
        text: "Invalid args: expected an object with 'entity' and 'operation' parameters"
      }]
    };
  }

  const { entity, operation, ...params } = args;

  if (!entity || !operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameters: entity and operation.\nEntities: plan, task, bucket, user\nPlan operations: list, get, create, update, delete\nTask operations: list, create, update, delete, get, assign, get_details, update_assignments, bulk_update, bulk_delete\nBucket operations: list, create, update, delete, get_tasks\nUser operations: lookup"
      }]
    };
  }

  const policyError = await enforcePlannerPolicy({ entity, operation, ...params });
  if (policyError) {
    return {
      content: [{
        type: "text",
        text: policyError
      }]
    };
  }

  try {
    let result;
    switch (entity) {
      case 'plan':
        result = await handlePlanOperations({ operation, ...params });
        break;

      case 'task':
        // Route to appropriate task handler based on operation
        switch (operation) {
          case 'list':
          case 'create':
          case 'update':
          case 'delete':
          case 'get':
          case 'assign':
            result = await handleTaskOperations({ operation, ...params });
            break;
          case 'get_details':
            result = await handleTaskDetails(params);
            break;
          case 'update_assignments':
            result = await handleTaskEnhanced({ operation: 'update_assignments', ...params });
            break;
          case 'bulk_update':
            result = await handleBulkOperations({ operation: 'update', ...params });
            break;
          case 'bulk_delete':
            result = await handleBulkOperations({ operation: 'delete', ...params });
            break;
          case 'get_assignments':
            result = await handleTaskAssignments({ operation: 'get', ...params });
            break;
          default:
            return {
              content: [{
                type: "text",
                text: `Invalid task operation: ${operation}. Valid: list, create, update, delete, get, assign, get_details, update_assignments, get_assignments, bulk_update, bulk_delete`
              }]
            };
        }
        break;

      case 'bucket':
        result = await handleBucketOperations({ operation, ...params });
        break;

      case 'user':
        result = await handleUserLookup(params);
        break;

      default:
        return {
          content: [{
            type: "text",
            text: `Invalid entity: ${entity}. Valid entities: plan, task, bucket, user`
          }]
        };
    }

    if (!['list', 'get', 'get_details', 'get_assignments', 'get_tasks', 'lookup'].includes(operation)) {
      await recordSideEffect('planner.write', 'success', params.taskId || params.planId || params.bucketId || null, {
        entity,
        operation
      });
    }

    return result;
  } catch (error) {
    if (!['list', 'get', 'get_details', 'get_assignments', 'get_tasks', 'lookup'].includes(operation)) {
      await recordSideEffect('planner.write', 'failed', params.taskId || params.planId || params.bucketId || null, {
        entity,
        operation,
        error: error.message
      });
    }
    throw error;
  }
}

// Export single consolidated tool
const plannerTools = [
  {
    name: 'planner',
    description: 'Unified Microsoft Planner management: plans, tasks, buckets, and user lookups. Use entity + operation routing.',
    inputSchema: {
      type: 'object',
      properties: {
        entity: {
          type: 'string',
          enum: ['plan', 'task', 'bucket', 'user'],
          description: 'The entity type to operate on'
        },
        operation: {
          type: 'string',
          description: 'Operation to perform. Plan: list/get/create/update/delete. Task: list/create/update/delete/get/assign/get_details/update_assignments/get_assignments/bulk_update/bulk_delete. Bucket: list/create/update/delete/get_tasks. User: lookup.'
        },
        // Plan parameters
        planId: { type: 'string', description: 'Plan ID' },
        title: { type: 'string', description: 'Plan/Task title' },
        groupId: { type: 'string', description: 'Microsoft 365 group ID (for plan create)' },
        // Task parameters
        taskId: { type: 'string', description: 'Task ID' },
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
          description: 'User ID(s) to unassign (for update_assignments)'
        },
        percentComplete: { type: 'number', description: 'Completion percentage (0-100)' },
        // Bucket parameters
        name: { type: 'string', description: 'Bucket name' },
        orderHint: { type: 'string', description: 'Order hint for bucket positioning' },
        // User parameters
        email: { type: 'string', description: 'Single user email address (for user lookup)' },
        emails: { type: 'array', items: { type: 'string' }, description: 'Multiple user email addresses (for user lookup)' },
        // Bulk parameters
        taskIds: { type: 'array', items: { type: 'string' }, description: 'Task IDs (for bulk operations)' },
        updates: {
          type: 'object',
          properties: {
            percentComplete: { type: 'number' },
            bucketId: { type: 'string' },
            dueDateTime: { type: 'string' }
          },
          description: 'Update data (for bulk_update)'
        },
        // Assignment parameters
        userId: { type: 'string', description: 'Single user ID (for assignment operations)' },
        userIds: { type: 'array', items: { type: 'string' }, description: 'Multiple user IDs (for assignment operations)' }
      },
      required: ['entity', 'operation']
    },
    handler: safeTool('planner', handlePlanner)
  }
];

module.exports = { plannerTools };
