/**
 * Microsoft To Do module
 * Provides access to To Do lists, tasks, and checklist items
 * Graph API: /me/todo/lists, /tasks, /checklistItems
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { safeTool } = require('../utils/errors');

async function handleTodo(args) {
  const { operation, ...params } = args || {};

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: list_lists, create_list, get_list, update_list, delete_list, list_tasks, create_task, get_task, update_task, delete_task, list_checklist, add_checklist_item, update_checklist_item"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      // ---- List operations ----
      case 'list_lists': {
        const response = await callGraphAPI(accessToken, 'GET', 'me/todo/lists', null, {
          $top: params.maxResults || 50,
          $select: 'id,displayName,isOwner,isShared,wellknownListName'
        });

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No To Do lists found." }] };
        }

        const listText = response.value.map(list => {
          const shared = list.isShared ? ' (shared)' : '';
          const wellKnown = list.wellknownListName && list.wellknownListName !== 'none'
            ? ` [${list.wellknownListName}]` : '';
          return `- ${list.displayName}${wellKnown}${shared}\n  ID: ${list.id}`;
        }).join('\n');

        return { content: [{ type: "text", text: `Found ${response.value.length} lists:\n\n${listText}` }] };
      }

      case 'create_list': {
        if (!params.displayName) {
          return { content: [{ type: "text", text: "Missing required parameter: displayName" }] };
        }

        const newList = await callGraphAPI(accessToken, 'POST', 'me/todo/lists', {
          displayName: params.displayName
        });

        return { content: [{ type: "text", text: `List created successfully!\nList ID: ${newList.id}\nName: ${newList.displayName}` }] };
      }

      case 'get_list': {
        if (!params.listId) {
          return { content: [{ type: "text", text: "Missing required parameter: listId" }] };
        }

        const list = await callGraphAPI(accessToken, 'GET', `me/todo/lists/${params.listId}`, null);

        return {
          content: [{
            type: "text",
            text: `List: ${list.displayName}\nID: ${list.id}\nOwner: ${list.isOwner}\nShared: ${list.isShared}`
          }]
        };
      }

      case 'update_list': {
        if (!params.listId || !params.displayName) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, displayName" }] };
        }

        await callGraphAPI(accessToken, 'PATCH', `me/todo/lists/${params.listId}`, {
          displayName: params.displayName
        });

        return { content: [{ type: "text", text: "List updated successfully!" }] };
      }

      case 'delete_list': {
        if (!params.listId) {
          return { content: [{ type: "text", text: "Missing required parameter: listId" }] };
        }

        await callGraphAPI(accessToken, 'DELETE', `me/todo/lists/${params.listId}`, null);

        return { content: [{ type: "text", text: "List deleted successfully!" }] };
      }

      // ---- Task operations ----
      case 'list_tasks': {
        if (!params.listId) {
          return { content: [{ type: "text", text: "Missing required parameter: listId" }] };
        }

        const queryParams = {
          $top: params.maxResults || 50,
          $select: 'id,title,status,importance,isReminderOn,dueDateTime,completedDateTime,body'
        };

        if (params.status) {
          queryParams.$filter = `status eq '${params.status}'`;
        }

        const response = await callGraphAPI(accessToken, 'GET',
          `me/todo/lists/${params.listId}/tasks`, null, queryParams
        );

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No tasks found in this list." }] };
        }

        const taskText = response.value.map(task => {
          const status = task.status === 'completed' ? '[x]' : '[ ]';
          const importance = task.importance === 'high' ? ' !' : '';
          const due = task.dueDateTime ? ` (due: ${task.dueDateTime.dateTime.split('T')[0]})` : '';
          return `${status} ${task.title}${importance}${due}\n    ID: ${task.id}`;
        }).join('\n');

        return { content: [{ type: "text", text: `Found ${response.value.length} tasks:\n\n${taskText}` }] };
      }

      case 'create_task': {
        if (!params.listId || !params.title) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, title" }] };
        }

        const taskData = { title: params.title };

        if (params.body) {
          taskData.body = { content: params.body, contentType: 'text' };
        }
        if (params.importance) {
          taskData.importance = params.importance;
        }
        if (params.dueDate) {
          taskData.dueDateTime = {
            dateTime: params.dueDate.includes('T') ? params.dueDate : `${params.dueDate}T00:00:00`,
            timeZone: 'UTC'
          };
        }
        if (params.reminderDateTime) {
          taskData.isReminderOn = true;
          taskData.reminderDateTime = {
            dateTime: params.reminderDateTime,
            timeZone: 'UTC'
          };
        }

        const newTask = await callGraphAPI(accessToken, 'POST',
          `me/todo/lists/${params.listId}/tasks`, taskData
        );

        return { content: [{ type: "text", text: `Task created!\nTask ID: ${newTask.id}\nTitle: ${newTask.title}` }] };
      }

      case 'get_task': {
        if (!params.listId || !params.taskId) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId" }] };
        }

        const task = await callGraphAPI(accessToken, 'GET',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}`, null
        );

        const info = [
          `Title: ${task.title}`,
          `Status: ${task.status}`,
          `Importance: ${task.importance}`,
          task.dueDateTime ? `Due: ${task.dueDateTime.dateTime}` : null,
          task.body?.content ? `Body: ${task.body.content}` : null,
          `ID: ${task.id}`
        ].filter(Boolean).join('\n');

        return { content: [{ type: "text", text: `Task Details:\n\n${info}` }] };
      }

      case 'update_task': {
        if (!params.listId || !params.taskId) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId" }] };
        }

        const updateData = {};
        if (params.title) updateData.title = params.title;
        if (params.status) updateData.status = params.status;
        if (params.importance) updateData.importance = params.importance;
        if (params.body) updateData.body = { content: params.body, contentType: 'text' };
        if (params.dueDate) {
          updateData.dueDateTime = {
            dateTime: params.dueDate.includes('T') ? params.dueDate : `${params.dueDate}T00:00:00`,
            timeZone: 'UTC'
          };
        }

        await callGraphAPI(accessToken, 'PATCH',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}`, updateData
        );

        return { content: [{ type: "text", text: "Task updated successfully!" }] };
      }

      case 'delete_task': {
        if (!params.listId || !params.taskId) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId" }] };
        }

        await callGraphAPI(accessToken, 'DELETE',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}`, null
        );

        return { content: [{ type: "text", text: "Task deleted successfully!" }] };
      }

      // ---- Checklist operations ----
      case 'list_checklist': {
        if (!params.listId || !params.taskId) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId" }] };
        }

        const response = await callGraphAPI(accessToken, 'GET',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}/checklistItems`, null
        );

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No checklist items found." }] };
        }

        const items = response.value.map(item => {
          const checked = item.isChecked ? '[x]' : '[ ]';
          return `${checked} ${item.displayName} (ID: ${item.id})`;
        }).join('\n');

        return { content: [{ type: "text", text: `Checklist items:\n\n${items}` }] };
      }

      case 'add_checklist_item': {
        if (!params.listId || !params.taskId || !params.displayName) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId, displayName" }] };
        }

        const item = await callGraphAPI(accessToken, 'POST',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}/checklistItems`,
          { displayName: params.displayName }
        );

        return { content: [{ type: "text", text: `Checklist item added!\nItem ID: ${item.id}` }] };
      }

      case 'update_checklist_item': {
        if (!params.listId || !params.taskId || !params.checklistItemId) {
          return { content: [{ type: "text", text: "Missing required parameters: listId, taskId, checklistItemId" }] };
        }

        const updateData = {};
        if (params.displayName) updateData.displayName = params.displayName;
        if (params.isChecked !== undefined) updateData.isChecked = params.isChecked;

        await callGraphAPI(accessToken, 'PATCH',
          `me/todo/lists/${params.listId}/tasks/${params.taskId}/checklistItems/${params.checklistItemId}`,
          updateData
        );

        return { content: [{ type: "text", text: "Checklist item updated!" }] };
      }

      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid: list_lists, create_list, get_list, update_list, delete_list, list_tasks, create_task, get_task, update_task, delete_task, list_checklist, add_checklist_item, update_checklist_item`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in todo ${operation}:`, error);
    return { content: [{ type: "text", text: `Error in todo ${operation}: ${error.message}` }] };
  }
}

const todoTools = [
  {
    name: 'todo',
    description: 'Microsoft To Do: manage lists, tasks, and checklist items',
    inputSchema: {
      type: 'object',
      properties: {
        operation: {
          type: 'string',
          enum: [
            'list_lists', 'create_list', 'get_list', 'update_list', 'delete_list',
            'list_tasks', 'create_task', 'get_task', 'update_task', 'delete_task',
            'list_checklist', 'add_checklist_item', 'update_checklist_item'
          ],
          description: 'Operation to perform'
        },
        listId: { type: 'string', description: 'To Do list ID' },
        taskId: { type: 'string', description: 'To Do task ID' },
        checklistItemId: { type: 'string', description: 'Checklist item ID' },
        title: { type: 'string', description: 'Task title' },
        displayName: { type: 'string', description: 'List or checklist item name' },
        body: { type: 'string', description: 'Task body content' },
        status: { type: 'string', enum: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'], description: 'Task status' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Task importance' },
        dueDate: { type: 'string', description: 'Due date (YYYY-MM-DD or ISO 8601)' },
        reminderDateTime: { type: 'string', description: 'Reminder date/time in ISO 8601' },
        isChecked: { type: 'boolean', description: 'Check/uncheck a checklist item' },
        maxResults: { type: 'number', description: 'Max results (default: 50)' }
      },
      required: ['operation']
    },
    handler: safeTool('todo', handleTodo)
  }
];

module.exports = { todoTools };
