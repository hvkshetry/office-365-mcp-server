/**
 * Consolidated Email module
 * Reduces from 11 tools to 4 tools with operation parameters
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Unified email handler for list, read, and send operations
 */
async function handleEmail(args) {
  console.error('handleEmail received raw args:', JSON.stringify(args));
  console.error('handleEmail args type:', typeof args);
  console.error('handleEmail args is null?', args === null);
  console.error('handleEmail args is undefined?', args === undefined);
  
  if (!args || typeof args !== 'object') {
    return {
      content: [{ 
        type: "text", 
        text: `DEBUG: Invalid args object. Type: ${typeof args}, Value: ${args}` 
      }]
    };
  }
  
  const { operation, ...params } = args;
  
  console.error('handleEmail after destructuring - operation:', operation);
  console.error('handleEmail after destructuring - params:', JSON.stringify(params));
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, read, send" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    console.error('handleEmail received args:', JSON.stringify(args));
    console.error('handleEmail operation:', operation);
    console.error('handleEmail params:', JSON.stringify(params));
    
    switch (operation) {
      case 'list':
        return await listEmails(accessToken, params);
      case 'read':
        return await readEmail(accessToken, params);
      case 'send':
        console.error('Calling sendEmail with params:', JSON.stringify(params));
        return await sendEmail(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, read, send` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email ${operation}:`, error);
    console.error('Error stack:', error.stack);
    return {
      content: [{ type: "text", text: `Error in email ${operation}: ${error.message}` }]
    };
  }
}

/**
 * Unified email search handler with different modes
 */
async function handleEmailSearch(args) {
  const { query, mode = 'basic', maxResults = 10, filterType, from, subject } = args;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (mode) {
      case 'basic':
        return await searchEmailsBasic(accessToken, { query, from, subject, maxResults });
      case 'enhanced':
        return await searchEmailsEnhanced(accessToken, { query, maxResults });
      case 'simple':
        return await searchEmailsSimple(accessToken, { query, filterType, maxResults });
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid mode: ${mode}. Valid modes are: basic, enhanced, simple` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email search (${mode}):`, error);
    return {
      content: [{ type: "text", text: `Error in email search: ${error.message}` }]
    };
  }
}

/**
 * Unified email move handler with batch capability
 */
async function handleEmailMove(args) {
  const { emailIds, destinationFolderId, batch = false } = args;
  
  if (!emailIds || !destinationFolderId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: emailIds and destinationFolderId" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    if (batch || emailIds.length > 5) {
      return await batchMoveEmails(accessToken, { emailIds, destinationFolderId });
    } else {
      return await moveEmails(accessToken, { emailIds, destinationFolderId });
    }
  } catch (error) {
    console.error('Error moving emails:', error);
    return {
      content: [{ type: "text", text: `Error moving emails: ${error.message}` }]
    };
  }
}

/**
 * Unified email folder handler
 */
async function handleEmailFolder(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listEmailFolders(accessToken);
      case 'create':
        return await createEmailFolder(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email folder ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in email folder: ${error.message}` }]
    };
  }
}

/**
 * Unified email rules handler
 */
async function handleEmailRules(args) {
  const { operation, enhanced = false, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return enhanced ? 
          await listEmailRulesEnhanced(accessToken) : 
          await listEmailRules(accessToken);
      case 'create':
        return enhanced ? 
          await createEmailRuleEnhanced(accessToken, params) : 
          await createEmailRule(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email rules ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in email rules: ${error.message}` }]
    };
  }
}

// Implementation functions (existing logic from original files)

async function listEmails(accessToken, params) {
  const { folderId, maxResults = 10 } = params;
  
  const endpoint = folderId ? 
    `me/mailFolders/${folderId}/messages` : 
    'me/messages';
  
  const queryParams = {
    $top: maxResults,
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No emails found." }]
    };
  }
  
  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    return `- ${email.subject}${attachments}\n  From: ${email.from.emailAddress.address}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails:\n\n${emailsList}` 
    }]
  };
}

async function readEmail(accessToken, params) {
  const { emailId } = params;
  
  if (!emailId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: emailId" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `me/messages/${emailId}`,
    null,
    {
      $select: 'subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments'
    }
  );
  
  let emailContent = `Subject: ${response.subject}\n`;
  emailContent += `From: ${response.from.emailAddress.name} <${response.from.emailAddress.address}>\n`;
  emailContent += `To: ${response.toRecipients.map(r => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(', ')}\n`;
  if (response.ccRecipients && response.ccRecipients.length > 0) {
    emailContent += `CC: ${response.ccRecipients.map(r => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(', ')}\n`;
  }
  emailContent += `Date: ${new Date(response.receivedDateTime).toLocaleString()}\n`;
  emailContent += `Has Attachments: ${response.hasAttachments ? 'Yes' : 'No'}\n\n`;
  emailContent += `Body:\n${response.body.content}`;
  
  return {
    content: [{ type: "text", text: emailContent }]
  };
}

async function sendEmail(accessToken, params) {
  try {
    console.error('sendEmail called with accessToken type:', typeof accessToken);
    console.error('sendEmail called with params type:', typeof params);
    console.error('sendEmail called with params value:', params);
    console.error('sendEmail called with params JSON:', JSON.stringify(params));
    
    // Basic validation
    if (!params || typeof params !== 'object') {
      console.error('DEBUG: Invalid params object detected');
      console.error('DEBUG: params is null?', params === null);
      console.error('DEBUG: params is undefined?', params === undefined);
      console.error('DEBUG: typeof params:', typeof params);
      return {
        content: [{ 
          type: "text", 
          text: `DEBUG: Invalid parameters object. Type: ${typeof params}, Value: ${params}` 
        }]
      };
    }
    
    // Check required parameters exist
    if (!params.to || !params.subject || !params.body) {
      return {
        content: [{ 
          type: "text", 
          text: "Missing required parameters: to, subject, and body" 
        }]
      };
    }
    
    // Ensure 'to' is always an array
    const toRecipients = Array.isArray(params.to) ? params.to : 
                         (typeof params.to === 'string' ? [params.to] : []);
    
    if (toRecipients.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "Invalid 'to' parameter. Please provide valid email address(es)." 
        }]
      };
    }
    
    // Create message object with proper structure
    const message = {
      subject: params.subject,
      body: {
        contentType: "HTML",
        content: params.body
      },
      toRecipients: toRecipients.map(email => ({
        emailAddress: { address: email }
      }))
    };
    
    // Add CC/BCC if they exist
    if (params.cc) {
      const ccRecipients = Array.isArray(params.cc) ? params.cc : [params.cc];
      if (ccRecipients.length > 0) {
        message.ccRecipients = ccRecipients.map(email => ({
          emailAddress: { address: email }
        }));
      }
    }
    
    if (params.bcc) {
      const bccRecipients = Array.isArray(params.bcc) ? params.bcc : [params.bcc];
      if (bccRecipients.length > 0) {
        message.bccRecipients = bccRecipients.map(email => ({
          emailAddress: { address: email }
        }));
      }
    }
    
    // Send the email with proper Microsoft Graph API format
    await callGraphAPI(
      accessToken,
      'POST',
      'me/sendMail',
      {
        message: message,
        saveToSentItems: true
      },
      null
    );
    
    return {
      content: [{ type: "text", text: "Email sent successfully!" }]
    };
  } catch (error) {
    console.error('Error in sendEmail:', error);
    console.error('Error stack:', error.stack);
    console.error('Params received:', JSON.stringify(params));
    return {
      content: [{ type: "text", text: `Email send error: ${error.message}` }]
    };
  }
}

async function searchEmailsBasic(accessToken, params) {
  const { query, from, subject, maxResults } = params;
  
  let searchQuery = query;
  if (from) searchQuery = `from:${from} AND ${searchQuery}`;
  if (subject) searchQuery = `subject:${subject} AND ${searchQuery}`;
  
  const queryParams = {
    $search: `"${searchQuery}"`,
    $top: maxResults,
    $select: config.EMAIL_SELECT_FIELDS
    // Note: $orderby cannot be used with $search - results are automatically sorted by sentDateTime
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/messages',
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No emails found matching your search." }]
    };
  }
  
  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    return `- ${email.subject}${attachments}\n  From: ${email.from.emailAddress.address}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails:\n\n${emailsList}` 
    }]
  };
}

async function searchEmailsEnhanced(accessToken, params) {
  const { query, maxResults } = params;
  
  const searchPayload = {
    requests: [
      {
        entityTypes: ["message"],
        query: {
          queryString: query
        },
        size: maxResults,
        fields: ["subject", "from", "receivedDateTime", "hasAttachments", "id"]
      }
    ]
  };
  
  try {
    const response = await callGraphAPI(
      accessToken,
      'POST',
      'search/query',
      searchPayload
    );
    
    const hits = response.value[0]?.hitsContainers[0]?.hits || [];
    
    if (hits.length === 0) {
      return {
        content: [{ type: "text", text: "No emails found matching your search." }]
      };
    }
    
    const emailsList = hits.map(hit => {
      const resource = hit.resource;
      const attachments = resource.hasAttachments ? ' 📎' : '';
      // Microsoft Search API may return ID in different formats
      const emailId = resource.id || hit.hitId || 'Not available';
      const fromAddress = resource.from?.emailAddress?.address || resource.from || 'Unknown sender';
      return `- ${resource.subject}${attachments}\n  From: ${fromAddress}\n  Date: ${new Date(resource.receivedDateTime).toLocaleString()}\n  ID: ${emailId}\n`;
    }).join('\n');
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${hits.length} emails using enhanced search:\n\n${emailsList}` 
      }]
    };
  } catch (error) {
    // Fall back to basic search if enhanced search fails
    return await searchEmailsBasic(accessToken, params);
  }
}

async function searchEmailsSimple(accessToken, params) {
  const { query, filterType = 'subject', maxResults } = params;
  
  let filterQuery;
  switch (filterType) {
    case 'from':
      filterQuery = `from/emailAddress/address eq '${query}'`;
      break;
    case 'body':
      filterQuery = `contains(body/content, '${query}')`;
      break;
    case 'subject':
    default:
      filterQuery = `contains(subject, '${query}')`;
      break;
  }
  
  const queryParams = {
    $filter: filterQuery,
    $top: maxResults,
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/messages',
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No emails found matching your search." }]
    };
  }
  
  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    return `- ${email.subject}${attachments}\n  From: ${email.from.emailAddress.address}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails:\n\n${emailsList}` 
    }]
  };
}

async function moveEmails(accessToken, params) {
  const { emailIds, destinationFolderId } = params;
  
  const results = [];
  
  for (const emailId of emailIds) {
    try {
      await callGraphAPI(
        accessToken,
        'POST',
        `me/messages/${emailId}/move`,
        { destinationId: destinationFolderId }
      );
      results.push({ emailId, status: 'success' });
    } catch (error) {
      results.push({ emailId, status: 'failed', error: error.message });
    }
  }
  
  const successCount = results.filter(r => r.status === 'success').length;
  const failureCount = results.filter(r => r.status === 'failed').length;
  
  let message = `Moved ${successCount} emails successfully.`;
  if (failureCount > 0) {
    message += ` ${failureCount} emails failed to move.`;
  }
  
  return {
    content: [{ type: "text", text: message }]
  };
}

async function batchMoveEmails(accessToken, params) {
  const { emailIds, destinationFolderId } = params;
  
  const batchSize = 20;
  const results = [];
  
  for (let i = 0; i < emailIds.length; i += batchSize) {
    const batch = emailIds.slice(i, i + batchSize);
    const requests = batch.map((emailId, index) => ({
      id: `${index}`,
      method: 'POST',
      url: `/me/messages/${emailId}/move`,
      body: { destinationId: destinationFolderId },
      headers: { 'Content-Type': 'application/json' }
    }));
    
    try {
      const batchResponse = await callGraphAPI(
        accessToken,
        'POST',
        '$batch',
        { requests }
      );
      
      batchResponse.responses.forEach((response, index) => {
        if (response.status === 200) {
          results.push({ emailId: batch[index], status: 'success' });
        } else {
          results.push({ 
            emailId: batch[index], 
            status: 'failed', 
            error: response.body?.error?.message || 'Unknown error' 
          });
        }
      });
    } catch (error) {
      batch.forEach(emailId => {
        results.push({ emailId, status: 'failed', error: error.message });
      });
    }
  }
  
  const successCount = results.filter(r => r.status === 'success').length;
  const failureCount = results.filter(r => r.status === 'failed').length;
  
  let message = `Batch moved ${successCount} emails successfully.`;
  if (failureCount > 0) {
    message += ` ${failureCount} emails failed to move.`;
  }
  
  return {
    content: [{ type: "text", text: message }]
  };
}

async function listEmailRules(accessToken) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders/inbox/messageRules'
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email rules found." }]
    };
  }
  
  const rulesList = response.value.map((rule, index) => {
    return `${index + 1}. ${rule.displayName}\n   Enabled: ${rule.isEnabled}\n   ID: ${rule.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} email rules:\n\n${rulesList}` 
    }]
  };
}

async function listEmailRulesEnhanced(accessToken) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders/inbox/messageRules'
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email rules found." }]
    };
  }
  
  const rulesList = response.value.map((rule, index) => {
    let details = `${index + 1}. ${rule.displayName}\n`;
    details += `   Enabled: ${rule.isEnabled}\n`;
    details += `   ID: ${rule.id}\n`;
    
    if (rule.conditions?.fromAddresses?.length > 0) {
      details += `   From: ${rule.conditions.fromAddresses.map(a => a.emailAddress.address).join(', ')}\n`;
    }
    
    if (rule.actions?.moveToFolder) {
      details += `   Action: Move to folder ${rule.actions.moveToFolder}\n`;
    }
    
    return details;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} email rules (enhanced view):\n\n${rulesList}` 
    }]
  };
}

async function createEmailRule(accessToken, params) {
  const { displayName, fromAddresses, moveToFolder, forwardTo } = params;
  
  if (!displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: displayName" 
      }]
    };
  }
  
  const rule = {
    displayName,
    sequence: 1,
    isEnabled: true,
    conditions: {},
    actions: {}
  };
  
  if (fromAddresses && fromAddresses.length > 0) {
    rule.conditions.fromAddresses = fromAddresses.map(email => ({
      emailAddress: { address: email }
    }));
  }
  
  if (moveToFolder) {
    rule.actions.moveToFolder = moveToFolder;
  }
  
  if (forwardTo && forwardTo.length > 0) {
    rule.actions.forwardTo = forwardTo.map(email => ({
      emailAddress: { address: email }
    }));
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/mailFolders/inbox/messageRules',
    rule
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Email rule created successfully!\nRule ID: ${response.id}` 
    }]
  };
}

async function createEmailRuleEnhanced(accessToken, params) {
  const { displayName, fromAddresses, moveToFolder, forwardTo, subjectContains, importance } = params;
  
  if (!displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: displayName" 
      }]
    };
  }
  
  const rule = {
    displayName,
    sequence: 1,
    isEnabled: true,
    conditions: {},
    actions: {}
  };
  
  if (fromAddresses && fromAddresses.length > 0) {
    rule.conditions.fromAddresses = fromAddresses.map(email => ({
      name: email,
      address: email
    }));
  }
  
  if (subjectContains && subjectContains.length > 0) {
    rule.conditions.subjectContains = subjectContains;
  }
  
  if (importance) {
    rule.conditions.importance = importance;
  }
  
  if (moveToFolder) {
    rule.actions.moveToFolder = moveToFolder;
  }
  
  if (forwardTo && forwardTo.length > 0) {
    rule.actions.forwardTo = forwardTo.map(email => ({
      emailAddress: { 
        name: email,
        address: email 
      }
    }));
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/mailFolders/inbox/messageRules',
    rule
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Enhanced email rule created successfully!\nRule ID: ${response.id}` 
    }]
  };
}

// Folder operation implementations
async function listEmailFolders(accessToken) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders',
    null,
    {
      $top: 100,
      $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount'
    }
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email folders found." }]
    };
  }
  
  const foldersList = response.value
    .filter(folder => folder.displayName !== 'Conversation History')
    .map((folder, index) => {
      return `${index + 1}. ${folder.displayName}\n   ID: ${folder.id}\n   Messages: ${folder.totalItemCount} (${folder.unreadItemCount} unread)\n`;
    }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} email folders:\n\n${foldersList}` 
    }]
  };
}

async function createEmailFolder(accessToken, params) {
  const { displayName, parentFolderId } = params;
  
  if (!displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: displayName" 
      }]
    };
  }
  
  const endpoint = parentFolderId
    ? `me/mailFolders/${parentFolderId}/childFolders`
    : 'me/mailFolders';
  
  const folderData = { displayName };
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    endpoint,
    folderData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Email folder created successfully!\nFolder ID: ${response.id}` 
    }]
  };
}

// Export consolidated tools
const emailTools = [
  {
    name: "email",
    description: "Manage emails: list, read, or send messages",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "read", "send"],
          description: "The operation to perform" 
        },
        // List parameters
        folderId: { type: "string", description: "Folder ID to list emails from (for list operation)" },
        maxResults: { type: "number", description: "Maximum number of results (default: 10)" },
        // Read parameters
        emailId: { type: "string", description: "Email ID to read (for read operation)" },
        // Send parameters
        to: { 
          type: "array", 
          items: { type: "string" },
          description: "Recipient email addresses (for send operation)" 
        },
        subject: { type: "string", description: "Email subject (for send operation)" },
        body: { type: "string", description: "Email body in HTML format (for send operation)" },
        cc: { 
          type: "array", 
          items: { type: "string" },
          description: "CC recipients (optional)" 
        },
        bcc: { 
          type: "array", 
          items: { type: "string" },
          description: "BCC recipients (optional)" 
        }
      },
      required: ["operation"]
    },
    handler: handleEmail
  },
  {
    name: "email_search",
    description: "Search emails with different modes: basic, enhanced, or simple",
    inputSchema: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query string" },
        mode: { 
          type: "string", 
          enum: ["basic", "enhanced", "simple"],
          description: "Search mode (default: basic)" 
        },
        maxResults: { type: "number", description: "Maximum number of results (default: 10)" },
        // Basic mode parameters
        from: { type: "string", description: "Filter by sender (basic mode)" },
        subject: { type: "string", description: "Filter by subject (basic mode)" },
        // Simple mode parameters
        filterType: { 
          type: "string", 
          enum: ["subject", "from", "body"],
          description: "Type of filter for simple mode (default: subject)" 
        }
      },
      required: ["query"]
    },
    handler: handleEmailSearch
  },
  {
    name: "email_move",
    description: "Move emails to a folder with optional batch processing",
    inputSchema: {
      type: "object",
      properties: {
        emailIds: { 
          type: "array", 
          items: { type: "string" },
          description: "Email IDs to move" 
        },
        destinationFolderId: { type: "string", description: "Destination folder ID" },
        batch: { type: "boolean", description: "Use batch processing for better performance (auto-enabled for >5 emails)" }
      },
      required: ["emailIds", "destinationFolderId"]
    },
    handler: handleEmailMove
  },
  {
    name: "email_folder",
    description: "Manage email folders: list or create",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "create"],
          description: "The operation to perform" 
        },
        // Create parameters
        displayName: { type: "string", description: "Folder display name (for create operation)" },
        parentFolderId: { type: "string", description: "Parent folder ID (optional, for creating subfolders)" }
      },
      required: ["operation"]
    },
    handler: handleEmailFolder
  },
  {
    name: "email_rules",
    description: "Manage email rules: list or create",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "create"],
          description: "The operation to perform" 
        },
        enhanced: { type: "boolean", description: "Use enhanced mode for more features (default: false)" },
        // Create parameters
        displayName: { type: "string", description: "Rule display name (for create operation)" },
        fromAddresses: { 
          type: "array", 
          items: { type: "string" },
          description: "Filter emails from these addresses" 
        },
        moveToFolder: { type: "string", description: "Folder ID to move emails to" },
        forwardTo: { 
          type: "array", 
          items: { type: "string" },
          description: "Email addresses to forward to" 
        },
        // Enhanced create parameters
        subjectContains: { 
          type: "array", 
          items: { type: "string" },
          description: "Filter by subject keywords (enhanced mode)" 
        },
        importance: { 
          type: "string", 
          enum: ["low", "normal", "high"],
          description: "Filter by importance (enhanced mode)" 
        }
      },
      required: ["operation"]
    },
    handler: handleEmailRules
  }
];

module.exports = { emailTools };
