/**
 * Microsoft Search API module - Unified Smart Search
 * Provides intelligent search across all Microsoft 365 services with advanced features
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Main search handler - single entry point for all search operations
 */
async function handleSearch(args) {
  try {
    const accessToken = await ensureAuthenticated();
    
    // Validate required parameters
    if (!args.query) {
      return {
        content: [{ 
          type: "text", 
          text: "Missing required parameter: query" 
        }]
      };
    }
    
    // Intelligent routing based on parameters
    if (args.siteId && !args.entityTypes) {
      // Site-specific search for lists and libraries
      return await searchSiteContent(accessToken, args);
    } else if (args.peopleSearch) {
      // Dedicated people search with user endpoints
      return await searchPeople(accessToken, args);
    } else {
      // Universal smart search with all Graph features
      return await smartUnifiedSearch(accessToken, args);
    }
  } catch (error) {
    console.error('Error in search:', error);
    return {
      content: [{ 
        type: "text", 
        text: `Search error: ${error.message}` 
      }]
    };
  }
}

/**
 * Smart unified search with all Graph API features enabled
 */
async function smartUnifiedSearch(accessToken, args) {
  const {
    query,
    entityTypes = ['driveItem', 'message', 'event', 'listItem'],
    limit = 25,
    from = 0,
    fileTypes,
    dateRange,
    aggregateBy = ['fileType', 'lastModifiedBy'],
    enrichContent = true,
    includeExcelData = true,
    sortBy,
    filters
  } = args;
  
  // Build intelligent KQL query
  const kqlQuery = buildSmartKQLQuery(query, {
    fileTypes,
    dateRange,
    filters,
    hasEmailFilters: query.includes('from:') || query.includes('to:') || query.includes('subject:')
  });
  
  // Validate and adjust entity types for compatibility
  const validEntityTypes = validateEntityTypes(entityTypes);
  
  // Build search request with all advanced features
  const searchRequest = {
    requests: [{
      entityTypes: validEntityTypes,
      query: {
        queryString: kqlQuery
      },
      from: from,
      size: Math.min(limit, 500), // Max 500 per Graph API limits
      
      // Always enable query improvements
      queryAlterationOptions: {
        enableSuggestion: true,
        enableModification: true
      },
      
      // Dynamic aggregations for faceted search (only for supported types)
      aggregations: shouldIncludeAggregations(validEntityTypes) 
        ? buildAggregations(aggregateBy, validEntityTypes)
        : undefined,
      
      // Rich field selection
      fields: [
        'id', 'name', 'webUrl', 'lastModifiedDateTime', 
        'size', 'createdBy', 'parentReference', 'file',
        'folder', 'package', 'specialFolder', 'root',
        'subject', 'from', 'to', 'receivedDateTime',
        'bodyPreview', 'hasAttachments', 'importance'
      ],
      
      // Collapse duplicate results (only for supported types)
      collapseProperties: shouldIncludeCollapseProperties(validEntityTypes) 
        ? [{
            fields: ['title'],
            limit: 1
          }]
        : undefined
    }]
  };
  
  // Add sorting if specified
  if (sortBy) {
    // Map common field names to Graph API field names
    const fieldMap = {
      'lastModifiedDateTime': 'lastModifiedTime',
      'createdDateTime': 'createdTime',
      'lastModified': 'lastModifiedTime',
      'created': 'createdTime',
      'modified': 'lastModifiedTime'
    };
    
    const sortField = fieldMap[sortBy.field] || sortBy.field || 'rank';
    
    searchRequest.requests[0].sortProperties = [{
      name: sortField,
      isDescending: sortBy.descending !== false
    }];
  }
  
  // Execute search
  const response = await callGraphAPI(
    accessToken,
    'POST',
    '/search/query',
    searchRequest
  );
  
  // Process and enrich results if requested
  const processedResults = enrichContent 
    ? await enrichSearchResults(accessToken, response, includeExcelData)
    : extractBasicResults(response);
    
  return formatSmartResults(processedResults);
}

/**
 * Check if aggregations should be included based on entity types
 */
function shouldIncludeAggregations(entityTypes) {
  // Aggregations are only supported for driveItem, listItem, and externalItem
  const aggregationSupportedTypes = ['driveItem', 'listItem', 'externalItem'];
  return entityTypes.some(type => aggregationSupportedTypes.includes(type));
}

/**
 * Check if collapse properties should be included based on entity types
 */
function shouldIncludeCollapseProperties(entityTypes) {
  // Collapse is only supported for file and externalItem entity types
  const collapseSupportedTypes = ['driveItem', 'listItem', 'externalItem'];
  return entityTypes.some(type => collapseSupportedTypes.includes(type));
}

/**
 * Build intelligent KQL query with enhancements
 */
function buildSmartKQLQuery(query, options) {
  let kql = query;
  const parts = [];
  
  // Add the base query if provided
  if (query && query.trim()) {
    parts.push(query);
  }
  
  // Auto-enhance with file type filters
  if (options.fileTypes && options.fileTypes.length > 0) {
    const typeFilter = options.fileTypes
      .map(t => `filetype:${t}`)
      .join(' OR ');
    parts.push(`(${typeFilter})`);
  }
  
  // Add date range intelligently
  if (options.dateRange) {
    // KQL date filtering uses simple date format and comparison operators
    // Convert ISO dates to simple format (YYYY-MM-DD)
    const formatDate = (isoDate) => {
      if (!isoDate) return null;
      // Extract just the date part from ISO string
      return isoDate.split('T')[0];
    };
    
    const startDate = formatDate(options.dateRange.start);
    const endDate = formatDate(options.dateRange.end);
    
    if (startDate && endDate) {
      // Use KQL range syntax with '..' operator
      parts.push(`LastModifiedTime:${startDate}..${endDate}`);
    } else if (startDate) {
      parts.push(`LastModifiedTime >= ${startDate}`);
    } else if (endDate) {
      parts.push(`LastModifiedTime <= ${endDate}`);
    }
  }
  
  // Add custom filters
  if (options.filters) {
    parts.push(options.filters);
  }
  
  // Combine all parts with AND operator
  kql = parts.filter(p => p).join(' AND ');
  
  return kql || '*'; // Return * for empty query to get all results
}

/**
 * Build dynamic aggregations based on requested fields
 */
function buildAggregations(fields, entityTypes) {
  const aggregationConfigs = {
    fileType: {
      field: 'fileType',
      size: 10,
      bucketDefinition: {
        sortBy: 'count',
        isDescending: true,
        minimumCount: 1
      }
    },
    lastModifiedBy: {
      field: 'lastModifiedBy',
      size: 5,
      bucketDefinition: {
        sortBy: 'count',
        isDescending: true,
        minimumCount: 1
      }
    },
    createdDateTime: {
      field: 'createdDateTime',
      size: 10,
      bucketDefinition: { 
        sortBy: 'keyAsString', 
        isDescending: true,
        minimumCount: 1,
        ranges: [
          { from: 'now-1d', to: 'now' },
          { from: 'now-7d', to: 'now-1d' },
          { from: 'now-30d', to: 'now-7d' },
          { from: 'now-365d', to: 'now-30d' }
        ]
      }
    },
    department: {
      field: 'department',
      size: 10,
      bucketDefinition: {
        sortBy: 'count',
        isDescending: true,
        minimumCount: 1
      }
    },
    author: {
      field: 'author',
      size: 10,
      bucketDefinition: {
        sortBy: 'count',
        isDescending: true,
        minimumCount: 1
      }
    }
  };
  
  // Filter aggregations based on entity types
  const applicableFields = fields.filter(field => {
    // Some aggregations only make sense for certain entity types
    if (field === 'fileType' && !entityTypes.some(t => ['driveItem', 'listItem'].includes(t))) {
      return false;
    }
    return true;
  });
  
  return applicableFields
    .map(f => aggregationConfigs[f])
    .filter(Boolean);
}

/**
 * Validate and adjust entity types for Graph API compatibility
 */
function validateEntityTypes(types) {
  const fileTypes = ['driveItem', 'listItem', 'list', 'drive'];
  const messageTypes = ['message', 'chatMessage'];
  const eventTypes = ['event'];
  const personTypes = ['person'];
  
  // Check for incompatible combinations
  const hasFileTypes = types.some(t => fileTypes.includes(t));
  const hasMessageTypes = types.some(t => messageTypes.includes(t));
  const hasEventTypes = types.some(t => eventTypes.includes(t));
  const hasPersonTypes = types.some(t => personTypes.includes(t));
  
  // If mixed incompatible types, prioritize based on first type
  if ((hasFileTypes && hasMessageTypes) || 
      (hasFileTypes && hasEventTypes) || 
      (hasMessageTypes && hasEventTypes) ||
      (hasPersonTypes && (hasFileTypes || hasMessageTypes || hasEventTypes))) {
    
    console.error('Warning: Incompatible entity types detected. Using compatible subset.');
    
    // Return compatible subset based on priority
    if (hasFileTypes) {
      return types.filter(t => fileTypes.includes(t));
    } else if (hasMessageTypes) {
      return types.filter(t => messageTypes.includes(t));
    } else if (hasEventTypes) {
      return types.filter(t => eventTypes.includes(t));
    } else if (hasPersonTypes) {
      return types.filter(t => personTypes.includes(t));
    }
  }
  
  return types;
}

/**
 * Enrich search results with additional content and metadata
 */
async function enrichSearchResults(accessToken, searchResponse, includeExcelData) {
  const hits = extractHits(searchResponse);
  const enrichedResults = [];
  
  // Process each result
  for (const hit of hits) {
    const enrichedHit = { ...hit };
    const resource = hit.resource;
    
    // Excel workbook enrichment
    if (includeExcelData && resource.name?.match(/\.xlsx?$/i) && resource.id) {
      try {
        const tables = await callGraphAPI(
          accessToken,
          'GET',
          `/me/drive/items/${resource.id}/workbook/tables`,
          null,
          { $top: 5 }
        );
        
        enrichedHit.structuredData = {
          type: 'excel',
          tableCount: tables.value.length,
          tables: tables.value.map(t => ({
            name: t.name,
            id: t.id,
            hasHeaders: t.showHeaders,
            rowCount: t.rows?.count || 0,
            columnCount: t.columns?.count || 0
          }))
        };
      } catch (error) {
        // Gracefully handle if workbook API fails
        console.error('Could not enrich Excel data:', error.message);
      }
    }
    
    // Add content preview for documents
    if (resource.file) {
      enrichedHit.preview = {
        mimeType: resource.file.mimeType,
        size: formatFileSize(resource.size),
        lastModified: resource.lastModifiedDateTime,
        quickActions: []
      };
      
      // Add quick action URLs for Office documents
      if (resource.name?.match(/\.(docx?|xlsx?|pptx?)$/i)) {
        enrichedHit.preview.quickActions.push({
          action: 'view',
          url: resource.webUrl
        });
      }
    }
    
    // Add email-specific enrichment
    if (hit.entityType === 'message' && resource.bodyPreview) {
      enrichedHit.emailContext = {
        preview: resource.bodyPreview.substring(0, 200),
        hasAttachments: resource.hasAttachments,
        importance: resource.importance,
        isRead: resource.isRead
      };
    }
    
    enrichedResults.push(enrichedHit);
  }
  
  // Extract query suggestions and aggregations
  const suggestions = searchResponse.value[0]?.queryAlterationResponse;
  const aggregations = searchResponse.value[0]?.aggregations || [];
  const totalCount = searchResponse.value[0]?.hitsContainers?.[0]?.total || enrichedResults.length;
  
  return {
    results: enrichedResults,
    suggestions,
    aggregations,
    totalCount
  };
}

/**
 * Extract basic results without enrichment
 */
function extractBasicResults(searchResponse) {
  const hits = extractHits(searchResponse);
  const suggestions = searchResponse.value[0]?.queryAlterationResponse;
  const aggregations = searchResponse.value[0]?.aggregations || [];
  const totalCount = searchResponse.value[0]?.hitsContainers?.[0]?.total || hits.length;
  
  return {
    results: hits,
    suggestions,
    aggregations,
    totalCount
  };
}

/**
 * Extract hits from search response
 */
function extractHits(response) {
  const hits = [];
  
  if (!response.value || !response.value[0]) {
    return hits;
  }
  
  for (const hitsContainer of response.value[0].hitsContainers || []) {
    const entityType = hitsContainer.entityType;
    const containerHits = hitsContainer.hits || [];
    
    for (const hit of containerHits) {
      hits.push({
        entityType,
        rank: hit.rank,
        summary: hit.summary,
        resource: hit.resource
      });
    }
  }
  
  return hits;
}

/**
 * Format smart search results with rich information
 */
function formatSmartResults(data) {
  const { results, suggestions, aggregations, totalCount } = data;
  
  let output = `Found ${totalCount} result${totalCount !== 1 ? 's' : ''}\n`;
  
  // Show query suggestions if available
  if (suggestions?.suggestion) {
    output += `\n💡 Did you mean: "${suggestions.suggestion.text}"?\n`;
    if (suggestions.modification) {
      output += `   Modified query: "${suggestions.modification.text}"\n`;
    }
  }
  
  // Show aggregations as filter options
  if (aggregations && aggregations.length > 0) {
    output += '\n📊 Filter by:\n';
    
    for (const agg of aggregations) {
      if (agg.buckets && agg.buckets.length > 0) {
        output += `  ${formatAggregationName(agg.field)}:\n`;
        
        const topBuckets = agg.buckets.slice(0, 5);
        for (const bucket of topBuckets) {
          const key = formatBucketKey(bucket.key, agg.field);
          output += `    • ${key} (${bucket.count})\n`;
        }
        
        if (agg.buckets.length > 5) {
          output += `    ... and ${agg.buckets.length - 5} more\n`;
        }
      }
    }
  }
  
  // Format results
  if (results.length > 0) {
    output += '\n📄 Results:\n';
    output += '─'.repeat(50) + '\n';
    
    results.forEach((result, idx) => {
      output += formatRichResult(result, idx + 1);
      if (idx < results.length - 1) {
        output += '─'.repeat(50) + '\n';
      }
    });
  } else {
    output += '\nNo results found. Try adjusting your search query.\n';
  }
  
  return {
    content: [{ type: "text", text: output }],
    metadata: { totalCount, aggregations, suggestions }
  };
}

/**
 * Format a single rich result
 */
function formatRichResult(result, index) {
  const resource = result.resource;
  let output = `\n${index}. `;
  
  // Title/Name
  const title = resource.name || resource.subject || resource.displayName || 'Untitled';
  output += `📌 ${title}\n`;
  
  // Entity type badge
  output += `   Type: ${formatEntityType(result.entityType)}\n`;
  
  // URL if available
  if (resource.webUrl) {
    output += `   🔗 ${resource.webUrl}\n`;
  }
  
  // File-specific information
  if (resource.file) {
    output += `   📁 ${result.preview?.mimeType || 'File'} | ${result.preview?.size || 'Unknown size'}\n`;
  }
  
  // Excel-specific structured data
  if (result.structuredData?.type === 'excel') {
    const data = result.structuredData;
    output += `   📊 Excel: ${data.tableCount} table${data.tableCount !== 1 ? 's' : ''}`;
    
    if (data.tables.length > 0) {
      output += ' (';
      output += data.tables.slice(0, 3).map(t => t.name).join(', ');
      if (data.tables.length > 3) {
        output += ', ...';
      }
      output += ')';
    }
    output += '\n';
  }
  
  // Email-specific information
  if (result.emailContext) {
    const email = result.emailContext;
    output += `   ✉️ ${email.importance === 'high' ? '❗ ' : ''}`;
    output += email.hasAttachments ? '📎 ' : '';
    output += `Preview: ${email.preview}...\n`;
  }
  
  // Message/Event specific
  if (resource.from?.emailAddress?.address) {
    output += `   From: ${resource.from.emailAddress.address}\n`;
  }
  if (resource.receivedDateTime) {
    output += `   Date: ${new Date(resource.receivedDateTime).toLocaleString()}\n`;
  }
  
  // Modified date for files
  if (resource.lastModifiedDateTime) {
    output += `   Modified: ${new Date(resource.lastModifiedDateTime).toLocaleString()}\n`;
  }
  
  // Created by
  if (resource.createdBy?.user?.displayName) {
    output += `   Created by: ${resource.createdBy.user.displayName}\n`;
  }
  
  // Summary if available
  if (result.summary) {
    output += `   Summary: ${result.summary.substring(0, 150)}...\n`;
  }
  
  return output + '\n';
}

/**
 * Search for people in the organization
 */
async function searchPeople(accessToken, args) {
  const { 
    query,
    limit = 10,
    select = 'displayName,mail,jobTitle,department,officeLocation,mobilePhone,userPrincipalName'
  } = args;
  
  // Build search query for people
  const queryParams = {
    $search: `"displayName:${query}" OR "mail:${query}" OR "jobTitle:${query}" OR "department:${query}"`,
    $select: select,
    $top: limit,
    $orderby: 'displayName'
  };
  
  // Add filter if provided
  if (args.filter) {
    queryParams.$filter = args.filter;
  }
  
  try {
    // People search requires ConsistencyLevel header
    const response = await callGraphAPI(
      accessToken,
      'GET',
      '/users',
      null,
      queryParams,
      { 'ConsistencyLevel': 'eventual' }
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No people found matching your search." }]
      };
    }
    
    // Format people results
    let output = `Found ${response.value.length} people:\n\n`;
    
    response.value.forEach((person, idx) => {
      output += `${idx + 1}. 👤 ${person.displayName}\n`;
      output += `   📧 ${person.mail || person.userPrincipalName || 'No email'}\n`;
      if (person.jobTitle) output += `   💼 ${person.jobTitle}\n`;
      if (person.department) output += `   🏢 ${person.department}\n`;
      if (person.officeLocation) output += `   📍 ${person.officeLocation}\n`;
      if (person.mobilePhone) output += `   📱 ${person.mobilePhone}\n`;
      output += '\n';
    });
    
    return {
      content: [{ type: "text", text: output }]
    };
  } catch (error) {
    console.error('Error searching people:', error);
    return {
      content: [{ 
        type: "text", 
        text: `Error searching people: ${error.message}` 
      }]
    };
  }
}

/**
 * Search within a specific SharePoint site
 */
async function searchSiteContent(accessToken, args) {
  const { 
    query,
    siteId,
    limit = 25
  } = args;
  
  try {
    // Search for lists in the site that match the query
    const queryParams = {
      $filter: `contains(displayName,'${query}') or contains(description,'${query}')`,
      $select: 'id,displayName,description,webUrl,createdDateTime,lastModifiedDateTime',
      $top: limit
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
        content: [{ type: "text", text: "No lists or libraries found in this site matching your search." }]
      };
    }
    
    // Format site content results
    let output = `Found ${response.value.length} lists/libraries in site:\n\n`;
    
    response.value.forEach((list, idx) => {
      output += `${idx + 1}. 📋 ${list.displayName}\n`;
      if (list.description) {
        output += `   📝 ${list.description}\n`;
      }
      output += `   🔗 ${list.webUrl}\n`;
      output += `   Modified: ${new Date(list.lastModifiedDateTime).toLocaleString()}\n`;
      output += `   ID: ${list.id}\n\n`;
    });
    
    return {
      content: [{ type: "text", text: output }]
    };
  } catch (error) {
    console.error('Error searching site content:', error);
    return {
      content: [{ 
        type: "text", 
        text: `Error searching site content: ${error.message}` 
      }]
    };
  }
}

/**
 * Helper function to format file size
 */
function formatFileSize(bytes) {
  if (!bytes) return 'Unknown size';
  
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  if (bytes === 0) return '0 Bytes';
  
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Helper function to format entity type
 */
function formatEntityType(type) {
  const typeMap = {
    'driveItem': '📁 File/Folder',
    'message': '✉️ Email',
    'event': '📅 Calendar Event',
    'listItem': '📋 List Item',
    'person': '👤 Person',
    'chatMessage': '💬 Chat Message'
  };
  
  return typeMap[type] || type;
}

/**
 * Helper function to format aggregation field names
 */
function formatAggregationName(field) {
  const nameMap = {
    'fileType': '📁 File Type',
    'lastModifiedBy': '✏️ Last Modified By',
    'createdDateTime': '📅 Created Date',
    'department': '🏢 Department',
    'author': '✍️ Author'
  };
  
  return nameMap[field] || field;
}

/**
 * Helper function to format bucket keys
 */
function formatBucketKey(key, field) {
  if (field === 'createdDateTime') {
    // Format date ranges
    if (key.includes('now-1d')) return 'Today';
    if (key.includes('now-7d')) return 'This Week';
    if (key.includes('now-30d')) return 'This Month';
    if (key.includes('now-365d')) return 'This Year';
    return key;
  }
  
  if (field === 'fileType') {
    // Format file types
    const typeMap = {
      'docx': 'Word',
      'xlsx': 'Excel',
      'pptx': 'PowerPoint',
      'pdf': 'PDF',
      'txt': 'Text',
      'jpg': 'Image (JPEG)',
      'png': 'Image (PNG)'
    };
    return typeMap[key] || key.toUpperCase();
  }
  
  return key;
}

// Export single unified search tool
const searchTools = [
  {
    name: "search",
    description: "Powerful unified search across Microsoft 365 with intelligent features, aggregations, and content enrichment",
    inputSchema: {
      type: "object",
      properties: {
        query: { 
          type: "string", 
          description: "Search query (supports KQL syntax and natural language)" 
        },
        
        // Optional targeting
        entityTypes: { 
          type: "array",
          items: { 
            type: "string",
            enum: ["driveItem", "listItem", "message", "event", "person", "chatMessage"]
          },
          description: "Types to search (default: driveItem, message, event, listItem)"
        },
        
        // Optional filters
        fileTypes: {
          type: "array",
          items: { type: "string" },
          description: "Filter by file extensions (e.g., ['docx', 'xlsx', 'pdf'])"
        },
        dateRange: {
          type: "object",
          properties: {
            start: { type: "string", description: "Start date (ISO format)" },
            end: { type: "string", description: "End date (ISO format)" }
          },
          description: "Date range filter"
        },
        filters: {
          type: "string",
          description: "Additional KQL filters"
        },
        
        // Search modes
        peopleSearch: {
          type: "boolean",
          description: "Enable people-specific search mode"
        },
        siteId: {
          type: "string",
          description: "Search within specific SharePoint site"
        },
        
        // Results configuration  
        limit: { 
          type: "number", 
          description: "Max results (default: 25, max: 500)" 
        },
        from: {
          type: "number",
          description: "Pagination offset (default: 0)"
        },
        sortBy: {
          type: "object",
          properties: {
            field: { type: "string", description: "Field to sort by" },
            descending: { type: "boolean", description: "Sort descending (default: true)" }
          },
          description: "Sort configuration"
        },
        
        // Advanced features (enabled by default)
        aggregateBy: {
          type: "array",
          items: { 
            type: "string",
            enum: ["fileType", "lastModifiedBy", "createdDateTime", "department", "author"]
          },
          description: "Fields to aggregate for filters (default: fileType, lastModifiedBy)"
        },
        enrichContent: {
          type: "boolean",
          description: "Extract additional content and metadata (default: true)"
        },
        includeExcelData: {
          type: "boolean",
          description: "Extract Excel workbook structure (default: true)"
        }
      },
      required: ["query"]
    },
    handler: handleSearch
  }
];

module.exports = { searchTools };