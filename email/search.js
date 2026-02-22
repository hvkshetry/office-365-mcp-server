/**
 * Email search - unified search with three-tier fallback strategy
 * Tier 1: Microsoft Search API (most powerful, relevance ranking)
 * Tier 2: Graph $search with KQL
 * Tier 3: $filter parameter (most reliable fallback)
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');
const { getFolderIdByName } = require('./folders');

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

  // Handle empty query string - skip $search and use $filter directly
  const isEmptyQuery = !query || query === "" || (typeof query === 'string' && query.trim() === "");

  try {
    const accessToken = await ensureAuthenticated();

    // Handle folder filtering
    const { mailbox } = args;
    let searchEndpoint = `${config.getMailboxPrefix(mailbox)}/messages`;
    let folderToSearch = null;

    if (folderId) {
      folderToSearch = folderId;
    } else if (folderName) {
      // Convert folder name to ID
      folderToSearch = await getFolderIdByName(accessToken, folderName, mailbox);
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
      searchEndpoint = `${config.getMailboxPrefix(mailbox)}/mailFolders/${folderToSearch}/messages`;
    }

    // If query is empty, go directly to $filter (for filtering by status, dates, etc.)
    if (isEmptyQuery) {
      console.error('Empty query detected, using $filter directly');
      return await searchUsingFilter(accessToken, {
        query: '',
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

    // Early exit to $filter for date-scoped searches ONLY if not KQL format
    if ((startDate || endDate) && !isKQLFormat(query || '')) {
      console.error('Date filters detected with non-KQL query, using $filter directly');
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

    // Three-tier execution strategy

    // Tier 1: Try Microsoft Search API (most powerful) - skip if folder is specified
    if ((useRelevance || isComplexKQLQuery(kqlQuery)) && !folderToSearch) {
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
      if (isKQLFormat(query || '')) {
        const simplifiedQuery = (query || '').replace(/\s+(OR|AND|NOT)\s+/gi, ' ').trim();
        return await searchUsingFilter(accessToken, {
          query: simplifiedQuery,
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
      } else {
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
    }
  } catch (error) {
    console.error('Error in unified email search:', error);
    return {
      content: [{ type: "text", text: `Error in email search: ${error.message}` }]
    };
  }
}

// ============== SEARCH HELPER FUNCTIONS ==============

/**
 * Build KQL query from parameters
 */
function buildKQLQuery(params) {
  let kql = [];

  if (!params.query || params.query === '' || params.query.trim() === '') {
    // Don't add any query term for empty queries
  } else if (isKQLFormat(params.query)) {
    kql.push(params.query);
  } else if (params.query !== '*') {
    kql.push(`(subject:"${params.query}" OR body:"${params.query}" OR from:"${params.query}")`);
  }

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

  const kqlPatterns = [
    /^OR\s+/i,
    /^AND\s+/i,
    /^NOT\s+/i,
    /\sOR\s+/i,
    /\sAND\s+/i,
    /\sNOT\s+/i
  ];

  return kqlOperators.some(op => query.includes(op)) ||
         kqlPatterns.some(pattern => pattern.test(query));
}

/**
 * Check if query is complex enough to warrant Microsoft Search API
 */
function isComplexKQLQuery(query) {
  const operatorCount = (query.match(/ AND | OR | NOT /g) || []).length;
  const hasDateRange = query.includes('received>=') || query.includes('received<=');
  const hasMultipleFilters = (query.match(/:/g) || []).length > 2;

  return operatorCount > 1 || hasDateRange || hasMultipleFilters;
}

/**
 * Parse relative date strings like '7d', '1w', '1m', '1y'
 */
function parseRelativeDate(dateStr) {
  if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
    return dateStr;
  }

  if (dateStr.match(/^\d+[dwmy]$/)) {
    const num = parseInt(dateStr);
    const unit = dateStr.slice(-1);
    const date = new Date();

    switch(unit) {
      case 'd':
        date.setDate(date.getDate() - num);
        break;
      case 'w':
        date.setDate(date.getDate() - (num * 7));
        break;
      case 'm':
        date.setMonth(date.getMonth() - num);
        break;
      case 'y':
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
  const { query, maxResults, useRelevance, includeDeleted, mailbox } = params;

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
  const { query, endpoint, maxResults, mailbox } = params;

  const queryParams = {
    $search: `"${query}"`,
    $top: Math.min(maxResults, 250),
    $select: config.EMAIL_SELECT_FIELDS
  };

  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    queryParams
  );

  if (!response.value || response.value.length === 0) {
    console.error('Graph $search returned 0 results, falling back to $filter');
    throw new Error('No results from $search, trigger fallback');
  }

  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    const importance = email.importance !== 'normal' ? ` [${email.importance}]` : '';
    const unread = !email.isRead ? ' *' : '';
    const fromAddress = email.from?.emailAddress?.address || email.from?.address || 'Unknown sender';

    return `- ${email.subject || '(No subject)'}${attachments}${importance}${unread}\n  From: ${fromAddress}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
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
  , mailbox } = params;

  let filters = [];

  if (query && query !== '' && query !== '*' && !from && !subject) {
    filters.push(`contains(subject, '${query}')`);
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
    filters.push(`hasAttachments eq ${hasAttachments ? 'true' : 'false'}`);
  }

  if (isRead !== undefined) {
    filters.push(`isRead eq ${isRead ? 'true' : 'false'}`);
  }

  if (importance) {
    filters.push(`importance eq '${importance}'`);
  }

  const dateField = 'receivedDateTime';

  if (startDate) {
    const date = parseRelativeDate(startDate);
    filters.push(`${dateField} ge ${date}T00:00:00Z`);
  }

  if (endDate) {
    const date = parseRelativeDate(endDate);
    filters.push(`${dateField} le ${date}T23:59:59Z`);
  }

  const filterQuery = filters.length > 0 ? filters.join(' and ') : null;

  // Graph API rejects $orderby combined with contains() in $filter
  // ("The restriction or sort order is too complex for this operation")
  const hasContainsFilter = filters.some(f => f.includes('contains('));
  const hasFolderSpecificEndpoint = endpoint && endpoint.includes('mailFolders');
  const shouldIncludeOrderBy = !hasContainsFilter && (!hasFolderSpecificEndpoint || filters.length === 0);

  const queryParams = {
    $top: Math.min(maxResults, 250),
    $select: config.EMAIL_SELECT_FIELDS
  };

  if (shouldIncludeOrderBy) {
    queryParams.$orderby = 'receivedDateTime desc';
  }

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
    const fromAddress = email.from?.emailAddress?.address || email.from?.address || 'Unknown sender';

    return `- ${email.subject || '(No subject)'}${attachments}${importance}${unread}\n  From: ${fromAddress}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails (using filter fallback):\n\n${emailsList}`
    }]
  };
}

// Deprecated search functions kept for internal compatibility
async function searchEmailsBasic(accessToken, params) {
  const { query, from, subject, maxResults, mailbox } = params;

  let searchQuery = query;
  if (from) searchQuery = `from:${from} AND ${searchQuery}`;
  if (subject) searchQuery = `subject:${subject} AND ${searchQuery}`;

  const queryParams = {
    $search: `"${searchQuery}"`,
    $top: maxResults,
    $select: config.EMAIL_SELECT_FIELDS
  };

  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/messages`,
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
    const fromAddress = email.from?.emailAddress?.address || email.from?.address || 'Unknown sender';
    return `- ${email.subject || '(No subject)'}${attachments}\n  From: ${fromAddress}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails:\n\n${emailsList}`
    }]
  };
}

async function searchEmailsEnhanced(accessToken, params) {
  const { query, maxResults, mailbox } = params;

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
    return await searchEmailsBasic(accessToken, params);
  }
}

async function searchEmailsSimple(accessToken, params) {
  const { query, filterType = 'subject', maxResults, mailbox } = params;

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
    `${config.getMailboxPrefix(mailbox)}/messages`,
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
    const fromAddress = email.from?.emailAddress?.address || email.from?.address || 'Unknown sender';
    return `- ${email.subject || '(No subject)'}${attachments}\n  From: ${fromAddress}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails:\n\n${emailsList}`
    }]
  };
}

module.exports = { handleEmailSearch };
