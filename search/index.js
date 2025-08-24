/**
 * Microsoft Search API module
 * Provides unified search across Microsoft 365 services
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Unified search handler for all search operations
 */
async function handleSearch(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: unified, people, sites, lists, messages" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'unified':
        return await unifiedSearch(accessToken, params);
      case 'people':
        return await searchPeople(accessToken, params);
      case 'sites':
        return await searchSites(accessToken, params);
      case 'lists':
        return await searchLists(accessToken, params);
      case 'messages':
        return await searchMessages(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: unified, people, sites, lists, messages` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in search ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in search ${operation}: ${error.message}` }]
    };
  }
}

/**
 * Unified search across all Microsoft 365
 */
async function unifiedSearch(accessToken, params) {
  const { 
    query, 
    entityTypes = ['driveItem', 'listItem', 'message', 'event', 'site'],
    from = 0,
    size = 25,
    fields,
    filters,
    sortBy,
    enableSpelling = true,
    enableQueryRules = true
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  // Validate and split entity types into compatible groups
  const fileTypes = ['driveItem', 'listItem', 'site', 'list', 'drive', 'externalItem'];
  const messageTypes = ['message', 'chatMessage'];
  const eventTypes = ['event'];
  const personTypes = ['person'];
  
  // Check which groups are requested
  const hasFileTypes = entityTypes.some(t => fileTypes.includes(t));
  const hasMessageTypes = entityTypes.some(t => messageTypes.includes(t));
  const hasEventTypes = entityTypes.some(t => eventTypes.includes(t));
  const hasPersonTypes = entityTypes.some(t => personTypes.includes(t));
  
  // If incompatible types are mixed, use only the first compatible group
  let validEntityTypes = entityTypes;
  if ((hasFileTypes && hasMessageTypes) || 
      (hasFileTypes && hasEventTypes) || 
      (hasMessageTypes && hasEventTypes) ||
      (hasPersonTypes && (hasFileTypes || hasMessageTypes || hasEventTypes))) {
    console.error('Warning: Incompatible entity types detected. Using only compatible types.');
    
    // Prioritize based on what's included
    if (hasFileTypes) {
      validEntityTypes = entityTypes.filter(t => fileTypes.includes(t));
    } else if (hasMessageTypes) {
      validEntityTypes = entityTypes.filter(t => messageTypes.includes(t));
    } else if (hasEventTypes) {
      validEntityTypes = entityTypes.filter(t => eventTypes.includes(t));
    } else if (hasPersonTypes) {
      validEntityTypes = entityTypes.filter(t => personTypes.includes(t));
    }
  }
  
  // Build search request
  const searchRequest = {
    requests: [{
      entityTypes: validEntityTypes,
      query: {
        queryString: query
      },
      from: from,
      size: size,
      enableSpelling: enableSpelling,
      includeQueryAlterationOptions: enableQueryRules
    }]
  };
  
  // Add fields selection if specified
  if (fields && fields.length > 0) {
    searchRequest.requests[0].fields = fields;
  }
  
  // Add filters if specified
  if (filters) {
    searchRequest.requests[0].query.queryFilter = filters;
  }
  
  // Add sorting if specified
  if (sortBy) {
    searchRequest.requests[0].sortProperties = [{
      name: sortBy.field,
      isDescending: sortBy.descending || false
    }];
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    '/search/query',
    searchRequest
  );
  
  // Process results
  const results = [];
  for (const hitsContainer of response.value[0].hitsContainers) {
    const entityType = hitsContainer.entityType;
    const hits = hitsContainer.hits || [];
    
    for (const hit of hits) {
      results.push({
        type: entityType,
        rank: hit.rank,
        summary: hit.summary,
        resource: hit.resource
      });
    }
  }
  
  if (results.length === 0) {
    return {
      content: [{ type: "text", text: "No results found for your search." }]
    };
  }
  
  // Format results by type
  const formattedResults = formatSearchResults(results);
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${results.length} results across Microsoft 365:\n\n${formattedResults}` 
    }]
  };
}

/**
 * Search for people in the organization
 */
async function searchPeople(accessToken, params) {
  const { 
    query, 
    top = 10,
    filter,
    select = 'displayName,mail,jobTitle,department,officeLocation,mobilePhone'
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  const queryParams = {
    $search: `"displayName:${query}" OR "mail:${query}" OR "jobTitle:${query}"`,
    $select: select,
    $top: top
  };
  
  if (filter) {
    queryParams.$filter = filter;
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    '/users',
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No people found matching your search." }]
    };
  }
  
  const peopleList = response.value.map((person, index) => {
    return `${index + 1}. ${person.displayName}
   Email: ${person.mail || 'N/A'}
   Title: ${person.jobTitle || 'N/A'}
   Department: ${person.department || 'N/A'}
   Office: ${person.officeLocation || 'N/A'}
   Mobile: ${person.mobilePhone || 'N/A'}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} people:\n\n${peopleList}` 
    }]
  };
}

/**
 * Search for SharePoint sites
 */
async function searchSites(accessToken, params) {
  const { 
    query,
    top = 10
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  const queryParams = {
    $search: query,
    $select: 'id,name,displayName,webUrl,description,createdDateTime',
    $top: top
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    '/sites',
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No sites found matching your search." }]
    };
  }
  
  const sitesList = response.value.map((site, index) => {
    return `${index + 1}. ${site.displayName || site.name}
   URL: ${site.webUrl}
   Description: ${site.description || 'N/A'}
   Created: ${new Date(site.createdDateTime).toLocaleString()}
   ID: ${site.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} sites:\n\n${sitesList}` 
    }]
  };
}

/**
 * Search for SharePoint lists
 */
async function searchLists(accessToken, params) {
  const { 
    query,
    siteId,
    top = 10
  } = params;
  
  if (!query || !siteId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: query and siteId" 
      }]
    };
  }
  
  const queryParams = {
    $filter: `contains(displayName,'${query}') or contains(description,'${query}')`,
    $select: 'id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime',
    $top: top
  };
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `/sites/${siteId}/lists`,
    null,
    queryParams
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No lists found matching your search." }]
    };
  }
  
  const listsList = response.value.map((list, index) => {
    return `${index + 1}. ${list.displayName}
   Description: ${list.description || 'N/A'}
   URL: ${list.webUrl}
   Modified: ${new Date(list.lastModifiedDateTime).toLocaleString()}
   ID: ${list.id}`;
  }).join('\n\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} lists:\n\n${listsList}` 
    }]
  };
}

/**
 * Enhanced message search with Microsoft Search API
 */
async function searchMessages(accessToken, params) {
  const { 
    query,
    from,
    to,
    subject,
    hasAttachments,
    startDate,
    endDate,
    size = 25,
    sortBy = 'relevance' // relevance, date
  } = params;
  
  if (!query) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: query" 
      }]
    };
  }
  
  // Build KQL query
  let kqlParts = [query];
  
  if (from) kqlParts.push(`from:${from}`);
  if (to) kqlParts.push(`to:${to}`);
  if (subject) kqlParts.push(`subject:"${subject}"`);
  if (hasAttachments !== undefined) kqlParts.push(`hasattachments:${hasAttachments}`);
  if (startDate) kqlParts.push(`received>=${startDate}`);
  if (endDate) kqlParts.push(`received<=${endDate}`);
  
  const kqlQuery = kqlParts.join(' AND ');
  
  const searchRequest = {
    requests: [{
      entityTypes: ['message'],
      query: {
        queryString: kqlQuery
      },
      size: size,
      enableSpelling: true
    }]
  };
  
  // Add sorting
  if (sortBy === 'date') {
    searchRequest.requests[0].sortProperties = [{
      name: 'receivedDateTime',
      isDescending: true
    }];
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    '/search/query',
    searchRequest
  );
  
  const hits = response.value[0].hitsContainers[0].hits || [];
  
  if (hits.length === 0) {
    return {
      content: [{ type: "text", text: "No messages found matching your search." }]
    };
  }
  
  const messagesList = hits.map((hit, index) => {
    const msg = hit.resource;
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
      text: `Found ${hits.length} messages:\n\n${messagesList}` 
    }]
  };
}

/**
 * Format search results by type
 */
function formatSearchResults(results) {
  const grouped = {};
  
  for (const result of results) {
    if (!grouped[result.type]) {
      grouped[result.type] = [];
    }
    grouped[result.type].push(result);
  }
  
  let formatted = '';
  
  for (const [type, items] of Object.entries(grouped)) {
    formatted += `\n=== ${type.toUpperCase()} (${items.length}) ===\n`;
    
    items.forEach((item, index) => {
      const resource = item.resource;
      
      switch (type) {
        case 'driveItem':
          formatted += `${index + 1}. ${resource.name}
   Path: ${resource.parentReference?.path || 'N/A'}
   Modified: ${new Date(resource.lastModifiedDateTime).toLocaleString()}\n\n`;
          break;
          
        case 'message':
          formatted += `${index + 1}. ${resource.subject}
   From: ${resource.from?.emailAddress?.address || 'N/A'}
   Date: ${new Date(resource.receivedDateTime).toLocaleString()}\n\n`;
          break;
          
        case 'event':
          formatted += `${index + 1}. ${resource.subject}
   Start: ${new Date(resource.start?.dateTime).toLocaleString()}
   Location: ${resource.location?.displayName || 'N/A'}\n\n`;
          break;
          
        case 'site':
          formatted += `${index + 1}. ${resource.displayName || resource.name}
   URL: ${resource.webUrl}
   Description: ${resource.description || 'N/A'}\n\n`;
          break;
          
        case 'listItem':
          formatted += `${index + 1}. ${resource.fields?.Title || 'Untitled'}
   Modified: ${resource.lastModifiedDateTime ? new Date(resource.lastModifiedDateTime).toLocaleString() : 'N/A'}\n\n`;
          break;
          
        default:
          formatted += `${index + 1}. ${JSON.stringify(resource).substring(0, 100)}...\n\n`;
      }
    });
  }
  
  return formatted;
}

// Export consolidated tool
const searchTools = [
  {
    name: "search",
    description: "Search across Microsoft 365: unified search, people, sites, lists, or messages",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["unified", "people", "sites", "lists", "messages"],
          description: "The type of search to perform" 
        },
        // Common parameters
        query: { type: "string", description: "Search query" },
        top: { type: "number", description: "Maximum number of results" },
        // Unified search parameters
        entityTypes: { 
          type: "array", 
          items: { 
            type: "string",
            enum: ["driveItem", "listItem", "message", "event", "site", "drive", "list"]
          },
          description: "Entity types to search" 
        },
        from: { type: "number", description: "Starting index for pagination" },
        size: { type: "number", description: "Number of results to return" },
        fields: { 
          type: "array", 
          items: { type: "string" },
          description: "Fields to return in results" 
        },
        filters: { type: "string", description: "Additional filters in KQL format" },
        sortBy: { 
          type: "object",
          properties: {
            field: { type: "string", description: "Field to sort by" },
            descending: { type: "boolean", description: "Sort in descending order" }
          },
          description: "Sort configuration" 
        },
        enableSpelling: { type: "boolean", description: "Enable spelling correction" },
        enableQueryRules: { type: "boolean", description: "Enable query rules and suggestions" },
        // People search parameters
        filter: { type: "string", description: "OData filter for people search" },
        select: { type: "string", description: "Fields to select for people" },
        // Sites/Lists parameters
        siteId: { type: "string", description: "Site ID for list search" },
        // Message search parameters
        to: { type: "string", description: "Filter by recipient" },
        subject: { type: "string", description: "Filter by subject" },
        hasAttachments: { type: "boolean", description: "Filter by attachment presence" },
        startDate: { type: "string", description: "Start date filter" },
        endDate: { type: "string", description: "End date filter" }
      },
      required: ["operation", "query"]
    },
    handler: handleSearch
  }
];

module.exports = { searchTools };