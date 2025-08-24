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
 * Unified email search handler - single powerful search with automatic optimization
 */
async function handleEmailSearch(args) {
  const { 
    query,           // Required: KQL or natural language
    from,            // Optional: sender email filter
    to,              // Optional: recipient email filter
    subject,         // Optional: subject line filter
    hasAttachments,  // Optional: boolean
    isRead,          // Optional: boolean
    importance,      // Optional: high/normal/low
    startDate,       // Optional: ISO or relative (7d/1w/1m/1y)
    endDate,         // Optional: ISO or relative
    folderId,        // Optional: folder ID to search in
    folderName,      // Optional: folder name (auto-converts to ID)
    maxResults = 25, // Optional: max 1000
    useRelevance = false, // Optional: relevance vs date sort
    includeDeleted = false // Optional: include deleted items
  } = args;
  
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
    
    // Handle folder filtering
    let searchEndpoint = 'me/messages';
    let folderToSearch = null;
    
    if (folderId) {
      folderToSearch = folderId;
    } else if (folderName) {
      // Convert folder name to ID
      folderToSearch = await getFolderIdByName(accessToken, folderName);
      if (!folderToSearch) {
        return {
          content: [{ 
            type: "text", 
            text: `Folder '${folderName}' not found. Check folder name or use folderId.` 
          }]
        };
      }
    }
    
    if (folderToSearch) {
      searchEndpoint = `me/mailFolders/${folderToSearch}/messages`;
    }
    
    // Build KQL query from parameters
    const kqlQuery = buildKQLQuery({
      query,
      from,
      to,
      subject,
      hasAttachments,
      isRead,
      importance,
      startDate,
      endDate
    });
    
    console.error(`Unified search - KQL Query: ${kqlQuery}`);
    console.error(`Search endpoint: ${searchEndpoint}`);
    
    // Three-tier execution strategy
    
    // Tier 1: Try Microsoft Search API (most powerful)
    if (useRelevance || isComplexKQLQuery(kqlQuery)) {
      try {
        return await searchUsingMicrosoftSearchAPI(accessToken, {
          query: kqlQuery,
          maxResults,
          useRelevance,
          includeDeleted
        });
      } catch (error) {
        console.error('Microsoft Search API failed, falling back:', error.message);
      }
    }
    
    // Tier 2: Try $search with KQL
    try {
      return await searchUsingGraphSearch(accessToken, {
        query: kqlQuery,
        endpoint: searchEndpoint,
        maxResults
      });
    } catch (error) {
      console.error('Graph $search failed, falling back to $filter:', error.message);
      
      // Tier 3: Fall back to $filter
      return await searchUsingFilter(accessToken, {
        query,
        from,
        to,
        subject,
        hasAttachments,
        isRead,
        importance,
        startDate,
        endDate,
        endpoint: searchEndpoint,
        maxResults
      });
    }
  } catch (error) {
    console.error('Error in unified email search:', error);
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

// ============== NEW UNIFIED SEARCH HELPER FUNCTIONS ==============

/**
 * Convert folder name to folder ID
 */
async function getFolderIdByName(accessToken, folderName) {
  // Check well-known folder names first
  const wellKnownFolders = {
    'inbox': 'inbox',
    'sent': 'sentitems',
    'sent items': 'sentitems', 
    'drafts': 'drafts',
    'deleted': 'deleteditems',
    'deleted items': 'deleteditems',
    'junk': 'junkemail',
    'junk email': 'junkemail',
    'archive': 'archive'
  };
  
  const lowerName = folderName.toLowerCase();
  if (wellKnownFolders[lowerName]) {
    return wellKnownFolders[lowerName];
  }
  
  // Search for custom folder by name
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { 
        $filter: `displayName eq '${folderName}'`,
        $select: 'id,displayName'
      }
    );
    
    if (response.value && response.value.length > 0) {
      return response.value[0].id;
    }
  } catch (error) {
    console.error(`Error finding folder by name: ${error.message}`);
  }
  
  return null;
}

/**
 * Build KQL query from parameters
 */
function buildKQLQuery(params) {
  let kql = [];
  
  // Check if query is already in KQL format
  if (params.query && isKQLFormat(params.query)) {
    kql.push(params.query);
  } else if (params.query) {
    // Convert natural language to KQL - search in subject, body, and from
    kql.push(`(subject:"${params.query}" OR body:"${params.query}" OR from:"${params.query}")`);
  }
  
  // Add filters
  if (params.from) {
    kql.push(`from:${params.from}`);
  }
  if (params.to) {
    kql.push(`to:${params.to}`);
  }
  if (params.subject && !isKQLFormat(params.query)) {
    kql.push(`subject:"${params.subject}"`);
  }
  if (params.hasAttachments !== undefined) {
    kql.push(`hasattachment:${params.hasAttachments}`);
  }
  if (params.isRead !== undefined) {
    kql.push(`isread:${params.isRead}`);
  }
  if (params.importance) {
    kql.push(`importance:${params.importance}`);
  }
  
  // Date ranges
  if (params.startDate) {
    const date = parseRelativeDate(params.startDate);
    kql.push(`received>=${date}`);
  }
  if (params.endDate) {
    const date = parseRelativeDate(params.endDate);
    kql.push(`received<=${date}`);
  }
  
  return kql.length > 0 ? kql.join(' AND ') : '';
}

/**
 * Check if query contains KQL operators
 */
function isKQLFormat(query) {
  // Check for KQL operators
  const kqlOperators = [
    ':',      // Property separator
    ' AND ',  // Boolean AND
    ' OR ',   // Boolean OR
    ' NOT ',  // Boolean NOT
    'from:',
    'to:',
    'subject:',
    'body:',
    'hasattachment:',
    'isread:',
    'importance:',
    'received:'
  ];
  
  return kqlOperators.some(op => query.includes(op));
}

/**
 * Check if query is complex enough to warrant Microsoft Search API
 */
function isComplexKQLQuery(query) {
  // Complex queries have multiple operators or date ranges
  const operatorCount = (query.match(/ AND | OR | NOT /g) || []).length;
  const hasDateRange = query.includes('received>=') || query.includes('received<=');
  const hasMultipleFilters = (query.match(/:/g) || []).length > 2;
  
  return operatorCount > 1 || hasDateRange || hasMultipleFilters;
}

/**
 * Parse relative date strings like '7d', '1w', '1m', '1y'
 */
function parseRelativeDate(dateStr) {
  // If already ISO format, return as-is
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
    return dateStr;
  }
  
  // Handle relative dates
  if (dateStr.match(/^\d+[dwmy]$/)) {
    const num = parseInt(dateStr);
    const unit = dateStr.slice(-1);
    const date = new Date();
    
    switch(unit) {
      case 'd': // days
        date.setDate(date.getDate() - num);
        break;
      case 'w': // weeks
        date.setDate(date.getDate() - (num * 7));
        break;
      case 'm': // months
        date.setMonth(date.getMonth() - num);
        break;
      case 'y': // years
        date.setFullYear(date.getFullYear() - num);
        break;
    }
    
    return date.toISOString().split('T')[0];
  }
  
  return dateStr;
}

/**
 * Search using Microsoft Search API for relevance ranking
 */
async function searchUsingMicrosoftSearchAPI(accessToken, params) {
  const { query, maxResults, useRelevance, includeDeleted } = params;
  
  const searchPayload = {
    requests: [
      {
        entityTypes: ["message"],
        query: {
          queryString: query
        },
        from: 0,
        size: Math.min(maxResults, 1000),
        fields: ["subject", "from", "toRecipients", "receivedDateTime", "hasAttachments", "id", "bodyPreview", "importance", "isRead"],
        enableTopResults: useRelevance
      }
    ]
  };
  
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
    const emailId = resource.id || hit.hitId || 'Not available';
    const fromAddress = resource.from?.emailAddress?.address || resource.from || 'Unknown sender';
    const importance = resource.importance ? ` [${resource.importance}]` : '';
    const unread = resource.isRead === false ? ' *' : '';
    
    return `- ${resource.subject}${attachments}${importance}${unread}\n  From: ${fromAddress}\n  Date: ${new Date(resource.receivedDateTime).toLocaleString()}\n  ID: ${emailId}\n`;
  }).join('\n');
  
  const sortNote = useRelevance ? ' (sorted by relevance)' : ' (sorted by date)';
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${hits.length} emails${sortNote}:\n\n${emailsList}` 
    }]
  };
}

/**
 * Search using Graph API $search parameter
 */
async function searchUsingGraphSearch(accessToken, params) {
  const { query, endpoint, maxResults } = params;
  
  const queryParams = {
    $search: `"${query}"`,
    $top: Math.min(maxResults, 250),
    $select: config.EMAIL_SELECT_FIELDS
    // Note: $orderby cannot be used with $search
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
      content: [{ type: "text", text: "No emails found matching your search." }]
    };
  }
  
  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    const importance = email.importance !== 'normal' ? ` [${email.importance}]` : '';
    const unread = !email.isRead ? ' *' : '';
    
    return `- ${email.subject}${attachments}${importance}${unread}\n  From: ${email.from.emailAddress.address}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails:\n\n${emailsList}` 
    }]
  };
}

/**
 * Search using $filter parameter (most reliable fallback)
 */
async function searchUsingFilter(accessToken, params) {
  const { 
    query, 
    from, 
    to, 
    subject, 
    hasAttachments, 
    isRead, 
    importance, 
    startDate, 
    endDate, 
    endpoint, 
    maxResults 
  } = params;
  
  let filters = [];
  
  // Build filter conditions
  if (query && !from && !subject) {
    // Simple text search in subject or body preview
    filters.push(`(contains(subject, '${query}') or contains(bodyPreview, '${query}'))`);
  }
  
  if (from) {
    filters.push(`from/emailAddress/address eq '${from}'`);
  }
  
  if (to) {
    filters.push(`toRecipients/any(r: r/emailAddress/address eq '${to}')`);
  }
  
  if (subject) {
    filters.push(`contains(subject, '${subject}')`);
  }
  
  if (hasAttachments !== undefined) {
    filters.push(`hasAttachments eq ${hasAttachments}`);
  }
  
  if (isRead !== undefined) {
    filters.push(`isRead eq ${isRead}`);
  }
  
  if (importance) {
    filters.push(`importance eq '${importance}'`);
  }
  
  if (startDate) {
    const date = parseRelativeDate(startDate);
    filters.push(`receivedDateTime ge ${date}T00:00:00Z`);
  }
  
  if (endDate) {
    const date = parseRelativeDate(endDate);
    filters.push(`receivedDateTime le ${date}T23:59:59Z`);
  }
  
  const filterQuery = filters.length > 0 ? filters.join(' and ') : null;
  
  const queryParams = {
    $top: Math.min(maxResults, 250),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  
  if (filterQuery) {
    queryParams.$filter = filterQuery;
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
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
    const importance = email.importance !== 'normal' ? ` [${email.importance}]` : '';
    const unread = !email.isRead ? ' *' : '';
    
    return `- ${email.subject}${attachments}${importance}${unread}\n  From: ${email.from.emailAddress.address}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails (using filter fallback):\n\n${emailsList}` 
    }]
  };
}

// ============== ORIGINAL SEARCH FUNCTIONS (DEPRECATED) ==============

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
        // Email move operations return 201 (Created) on success
        if (response.status === 200 || response.status === 201) {
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

/**
 * Get focused inbox messages
 */
async function getFocusedInbox(accessToken, params) {
  const { maxResults = 25 } = params;
  
  const queryParams = {
    $filter: 'inferenceClassification eq \'Focused\'',
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime DESC',
    $top: maxResults
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
      content: [{ type: "text", text: "No focused messages found." }]
    };
  }
  
  const messagesList = response.value.map((msg, index) => {
    return `${index + 1}. ${msg.subject}
   From: ${msg.from?.emailAddress?.address || 'N/A'}
   Date: ${new Date(msg.receivedDateTime).toLocaleString()}
   Preview: ${msg.bodyPreview?.substring(0, 100)}...
   Has Attachments: ${msg.hasAttachments ? 'Yes' : 'No'}
   ID: ${msg.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} focused messages:\n\n${messagesList}` 
    }]
  };
}

/**
 * Manage email categories
 */
async function handleEmailCategories(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create, update, delete, apply, remove" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listCategories(accessToken);
      case 'create':
        return await createCategory(accessToken, params);
      case 'update':
        return await updateCategory(accessToken, params);
      case 'delete':
        return await deleteCategory(accessToken, params);
      case 'apply':
        return await applyCategory(accessToken, params);
      case 'remove':
        return await removeCategory(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in categories ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error: ${error.message}` }]
    };
  }
}

async function listCategories(accessToken) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/outlook/masterCategories'
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No categories found." }]
    };
  }
  
  const categoriesList = response.value.map((cat, index) => {
    return `${index + 1}. ${cat.displayName} (Color: ${cat.color})`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Categories:\n${categoriesList}` 
    }]
  };
}

async function createCategory(accessToken, params) {
  const { displayName, color = 'preset0' } = params;
  
  if (!displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: displayName" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/outlook/masterCategories',
    { displayName, color }
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Category '${displayName}' created with color ${color}` 
    }]
  };
}

async function applyCategory(accessToken, params) {
  const { emailId, categories } = params;
  
  if (!emailId || !categories) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: emailId and categories" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `me/messages/${emailId}`,
    { categories: Array.isArray(categories) ? categories : [categories] }
  );
  
  return {
    content: [{ type: "text", text: "Categories applied successfully!" }]
  };
}

/**
 * Get mail tips for recipients
 */
async function getMailTips(accessToken, params) {
  const { recipients } = params;
  
  if (!recipients || recipients.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: recipients" 
      }]
    };
  }
  
  const body = {
    EmailAddresses: recipients.map(email => ({ address: email })),
    MailTipsOptions: 'automaticReplies, mailboxFullStatus, customMailTip, deliveryRestriction, moderationStatus'
  };
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/getMailTips',
    body
  );
  
  const tips = response.value.map((tip, index) => {
    let tipInfo = `${recipients[index]}:\n`;
    
    if (tip.automaticReplies?.message) {
      tipInfo += `  - Auto-reply: ${tip.automaticReplies.message}\n`;
    }
    if (tip.mailboxFull) {
      tipInfo += `  - Mailbox is full\n`;
    }
    if (tip.customMailTip) {
      tipInfo += `  - Custom tip: ${tip.customMailTip}\n`;
    }
    if (tip.deliveryRestricted) {
      tipInfo += `  - Delivery restricted\n`;
    }
    if (tip.isModerated) {
      tipInfo += `  - Messages are moderated\n`;
    }
    
    return tipInfo;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Mail Tips:\n${tips}` 
    }]
  };
}

/**
 * Handle email with mentions
 */
async function handleMentions(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, get" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listMentions(accessToken, params);
      case 'get':
        return await getMentions(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in mentions ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error: ${error.message}` }]
    };
  }
}

async function listMentions(accessToken, params) {
  const { maxResults = 25 } = params;
  
  const queryParams = {
    $filter: 'mentionsPreview/isMentioned eq true',
    $select: 'id,subject,from,receivedDateTime,bodyPreview,mentionsPreview',
    $orderby: 'receivedDateTime DESC',
    $top: maxResults
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
      content: [{ type: "text", text: "No messages with mentions found." }]
    };
  }
  
  const messagesList = response.value.map((msg, index) => {
    return `${index + 1}. ${msg.subject}
   From: ${msg.from?.emailAddress?.address || 'N/A'}
   Date: ${new Date(msg.receivedDateTime).toLocaleString()}
   ID: ${msg.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} messages with mentions:\n\n${messagesList}` 
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
    description: "Unified email search with KQL support, folder filtering, and automatic optimization",
    inputSchema: {
      type: "object",
      properties: {
        query: { 
          type: "string", 
          description: "Search text or KQL syntax (e.g., 'project' or 'from:john@example.com AND subject:report')" 
        },
        from: { 
          type: "string", 
          description: "Filter by sender email address" 
        },
        to: { 
          type: "string", 
          description: "Filter by recipient email address" 
        },
        subject: { 
          type: "string", 
          description: "Filter by subject line" 
        },
        hasAttachments: { 
          type: "boolean", 
          description: "Filter emails with/without attachments" 
        },
        isRead: { 
          type: "boolean", 
          description: "Filter by read/unread status" 
        },
        importance: { 
          type: "string", 
          enum: ["high", "normal", "low"],
          description: "Filter by importance level" 
        },
        startDate: { 
          type: "string", 
          description: "Start date - ISO format (2025-08-01) or relative (7d/1w/1m/1y)" 
        },
        endDate: { 
          type: "string", 
          description: "End date - ISO format or relative" 
        },
        folderId: { 
          type: "string",
          description: "Specific folder ID to search in"
        },
        folderName: { 
          type: "string",
          description: "Folder name (inbox/sent/drafts/deleted/junk/archive or custom name)"
        },
        maxResults: { 
          type: "number", 
          description: "Max results 1-1000 (default: 25)" 
        },
        useRelevance: { 
          type: "boolean", 
          description: "Sort by relevance instead of date (uses Microsoft Search API)" 
        },
        includeDeleted: { 
          type: "boolean", 
          description: "Include deleted items in search results" 
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
,
  {
    name: "email_focused",
    description: "Get focused inbox messages (important messages filtered by AI)",
    inputSchema: {
      type: "object",
      properties: {
        maxResults: { type: "number", description: "Maximum number of results (default: 25)" }
      }
    },
    handler: async (args) => {
      const accessToken = await ensureAuthenticated();
      return await getFocusedInbox(accessToken, args);
    }
  },
  {
    name: "email_categories",
    description: "Manage email categories: list, create, update, delete, apply, or remove",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "create", "update", "delete", "apply", "remove"],
          description: "The operation to perform" 
        },
        // Create/Update parameters
        displayName: { type: "string", description: "Category display name" },
        color: { 
          type: "string",
          description: "Category color (preset0-preset24)" 
        },
        // Apply/Remove parameters
        emailId: { type: "string", description: "Email ID to apply/remove category" },
        categories: { 
          type: "array",
          items: { type: "string" },
          description: "Categories to apply/remove" 
        },
        // Update/Delete parameters
        categoryId: { type: "string", description: "Category ID for update/delete" }
      },
      required: ["operation"]
    },
    handler: handleEmailCategories
  }
  // Removed email_mailtips and email_mentions - not functional with current permissions/setup
];

module.exports = { emailTools };