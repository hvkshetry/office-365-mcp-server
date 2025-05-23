/**
 * Subscription management functionality
 */
const config = require('../config');
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Create subscription handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateSubscription(args) {
  const { resource, changeType, notificationUrl, expirationMinutes } = args;
  
  if (!resource || !changeType || !notificationUrl) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide resource, changeType, and notificationUrl." 
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Calculate expiration time (default to 60 minutes if not specified)
    const minutes = expirationMinutes || 60;
    const expirationDateTime = new Date();
    expirationDateTime.setMinutes(expirationDateTime.getMinutes() + minutes);
    
    // Prepare the subscription payload
    const subscriptionPayload = {
      changeType: changeType,
      notificationUrl: notificationUrl,
      resource: resource,
      expirationDateTime: expirationDateTime.toISOString(),
      clientState: "secretClientState"
    };
    
    // Make API call
    const response = await callGraphAPI(
      accessToken,
      'POST',
      '/subscriptions',
      subscriptionPayload
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Subscription created successfully!\n\nID: ${response.id}\nResource: ${response.resource}\nExpires: ${response.expirationDateTime}` 
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error creating subscription: ${error.message}` 
      }]
    };
  }
}

/**
 * List subscriptions handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSubscriptions(args) {
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Make API call
    const response = await callGraphAPI(accessToken, 'GET', '/subscriptions', null);
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No active subscriptions found."
        }]
      };
    }
    
    // Format results
    const subscriptionsList = response.value.map((subscription, index) => {
      const expiresAt = new Date(subscription.expirationDateTime).toLocaleString();
      return `${index + 1}. ID: ${subscription.id}\n   Resource: ${subscription.resource}\n   Change Type: ${subscription.changeType}\n   Expires: ${expiresAt}\n`;
    }).join("\n");
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} active subscriptions:\n\n${subscriptionsList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error listing subscriptions: ${error.message}`
      }]
    };
  }
}

/**
 * Renew subscription handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleRenewSubscription(args) {
  const { subscriptionId, expirationMinutes } = args;
  
  if (!subscriptionId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: subscriptionId"
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Calculate new expiration time
    const minutes = expirationMinutes || 60;
    const expirationDateTime = new Date();
    expirationDateTime.setMinutes(expirationDateTime.getMinutes() + minutes);
    
    // Prepare the update payload
    const updatePayload = {
      expirationDateTime: expirationDateTime.toISOString()
    };
    
    // Make API call
    await callGraphAPI(
      accessToken,
      'PATCH',
      `/subscriptions/${subscriptionId}`,
      updatePayload
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Subscription renewed successfully! New expiration time: ${expirationDateTime.toLocaleString()}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error renewing subscription: ${error.message}`
      }]
    };
  }
}

/**
 * Delete subscription handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteSubscription(args) {
  const { subscriptionId } = args;
  
  if (!subscriptionId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: subscriptionId"
      }]
    };
  }
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Make API call
    await callGraphAPI(
      accessToken,
      'DELETE',
      `/subscriptions/${subscriptionId}`,
      null
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Subscription deleted successfully!`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error deleting subscription: ${error.message}`
      }]
    };
  }
}

// Tool definitions
const subscriptionTools = [
  {
    name: "notification_create_subscription",
    description: "Creates a new webhook subscription for change notifications",
    inputSchema: {
      type: "object",
      properties: {
        resource: {
          type: "string",
          description: "The resource to subscribe to changes for"
        },
        changeType: {
          type: "string",
          description: "The type of changes to receive notifications for (created, updated, deleted)"
        },
        notificationUrl: {
          type: "string",
          description: "The URL where notifications should be sent"
        },
        expirationMinutes: {
          type: "number",
          description: "The number of minutes until the subscription expires (default: 60, max: 4230)"
        }
      },
      required: ["resource", "changeType", "notificationUrl"]
    },
    handler: handleCreateSubscription
  },
  {
    name: "notification_list_subscriptions",
    description: "Lists all active webhook subscriptions",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListSubscriptions
  },
  {
    name: "notification_renew_subscription",
    description: "Renews an existing webhook subscription",
    inputSchema: {
      type: "object",
      properties: {
        subscriptionId: {
          type: "string",
          description: "The ID of the subscription to renew"
        },
        expirationMinutes: {
          type: "number",
          description: "The number of minutes until the subscription expires (default: 60, max: 4230)"
        }
      },
      required: ["subscriptionId"]
    },
    handler: handleRenewSubscription
  },
  {
    name: "notification_delete_subscription",
    description: "Deletes an existing webhook subscription",
    inputSchema: {
      type: "object",
      properties: {
        subscriptionId: {
          type: "string",
          description: "The ID of the subscription to delete"
        }
      },
      required: ["subscriptionId"]
    },
    handler: handleDeleteSubscription
  }
];

module.exports = {
  subscriptionTools,
  handleCreateSubscription,
  handleListSubscriptions,
  handleRenewSubscription,
  handleDeleteSubscription
};
