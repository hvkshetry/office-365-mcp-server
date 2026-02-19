/**
 * Consolidated Notifications module
 * Merges 4 tools → 1 tool with operation-based routing
 */
const {
  handleCreateSubscription,
  handleListSubscriptions,
  handleRenewSubscription,
  handleDeleteSubscription
} = require('./subscriptions');
const { safeTool } = require('../utils/errors');

/**
 * Unified notification handler
 */
async function handleNotification(args) {
  const { operation, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: create, list, renew, delete"
      }]
    };
  }

  switch (operation) {
    case 'create':
      return await handleCreateSubscription(params);
    case 'list':
      return await handleListSubscriptions(params);
    case 'renew':
      return await handleRenewSubscription(params);
    case 'delete':
      return await handleDeleteSubscription(params);
    default:
      return {
        content: [{
          type: "text",
          text: `Invalid operation: ${operation}. Valid operations: create, list, renew, delete`
        }]
      };
  }
}

const notificationTools = [
  {
    name: "notifications",
    description: "Manage webhook subscriptions for change notifications: create, list, renew, or delete",
    inputSchema: {
      type: "object",
      properties: {
        operation: {
          type: "string",
          enum: ["create", "list", "renew", "delete"],
          description: "The operation to perform"
        },
        resource: {
          type: "string",
          description: "The resource to subscribe to changes for (for create)"
        },
        changeType: {
          type: "string",
          description: "The type of changes: created, updated, deleted (for create)"
        },
        notificationUrl: {
          type: "string",
          description: "The URL where notifications should be sent (for create)"
        },
        expirationMinutes: {
          type: "number",
          description: "Minutes until subscription expires (default: 60, max: 4230)"
        },
        subscriptionId: {
          type: "string",
          description: "Subscription ID (for renew/delete)"
        }
      },
      required: ["operation"]
    },
    handler: safeTool('notifications', handleNotification)
  }
];

module.exports = { notificationTools };
