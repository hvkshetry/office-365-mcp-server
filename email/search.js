/**
 * Email search - unified search with three-tier fallback strategy
 * Tier 1: Microsoft Search API (most powerful, relevance ranking)
 * Tier 2: Graph $search with KQL
 * Tier 3: $filter parameter (most reliable fallback)
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { validateId } = require('../utils/validate');
const config = require('../config');
const { getFolderIdByName } = require('./folders');
const { withMailStructuredContent } = require('./results');

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
  // from/to/hasAttachments/importance cannot be expressed in a Graph message $filter without a
  // "restriction or sort order is too complex" 400 (from/to also fail the filter node itself);
  // they must ride the KQL $search path, which Graph returns already date-sorted.
  const usesKqlOnly = !!(from || to || hasAttachments !== undefined || importance);

  try {
    const accessToken = await ensureAuthenticated();

    // Handle folder filtering
    const { mailbox } = args;
    let searchEndpoint = `${config.getMailboxPrefix(mailbox)}/messages`;
    let folderToSearch = null;

    if (folderId) {
      validateId(folderId, 'folderId');
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

    // If query is empty, route by predicate type: from/to/hasAttachments/importance go through
    // $search (Graph rejects them in $filter); pure date/status stay on the $filter path.
    if (isEmptyQuery) {
      if (usesKqlOnly) {
        const kqlOnly = buildKQLQuery({ query: '', from, to, subject, hasAttachments, isRead, importance, startDate, endDate });
        console.error(`Empty query with $search-only predicate(s), using $search: ${kqlOnly}`);
        try {
          return await searchUsingGraphSearch(accessToken, {
            query: kqlOnly,
            endpoint: searchEndpoint,
            maxResults,
            allowEmpty: true
          });
        } catch (error) {
          // There is no second tier for these predicates — Graph rejects from/to/hasAttachments/
          // importance in a $filter. Say that, rather than letting a transient 503 surface as a
          // bare error the caller cannot act on or distinguish from an empty mailbox.
          console.error('Graph $search failed on a $search-only predicate set:', error.message);
          return { isError: true, content: [{ type: "text", text:
            `Email search failed (${error.message}). from/to/hasAttachments/importance `
            + `cannot fall back to a subject filter, so no second attempt was made. This is NOT `
            + `evidence the mailbox is empty; retry, or search by a distinctive keyword instead.` }] };
        }
      }
      console.error('Empty query detected, using $filter directly');
      return await searchUsingFilter(accessToken, {
        query: '',
        subject,
        isRead,
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

    // Dates belong in the KQL expression too. Routing text + dates directly to
    // $filter silently discards body matches before a full-text search is tried.

    // Three-tier execution strategy
    // External mailbox: skip Tier 1 (/search/query) which doesn't scope to other users
    // and returns immutable IDs incompatible with read/get_attachment across mailboxes.
    // Tier 2 ($search on the scoped REST endpoint) handles external mailboxes correctly.

    // Tier 1: Try Microsoft Search API (most powerful) - skip if folder or external mailbox
    if ((useRelevance || isComplexKQLQuery(kqlQuery)) && !folderToSearch && (!mailbox || mailbox === 'me')) {
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
        maxResults,
        allowEmpty: usesKqlOnly
      });
    } catch (error) {
      console.error('Graph $search failed, falling back to $filter:', error.message);

      // Why Tier 2 gave up, carried into the Tier 3 result text. Silently degrading from a
      // body-capable search to a subject-only one produces an answer that LOOKS authoritative
      // and is not: a caller reads "no emails found" as "the mailbox has nothing", when in
      // fact bodies were never read. A transient 401/429/5xx must not be indistinguishable
      // from a genuine empty mailbox.
      const degradedBecause = error && error.isEmptyResult
        ? '$search matched nothing'
        : `$search failed (${(error && error.message) || 'unknown error'})`;

      // from/to/hasAttachments/importance cannot fall back to $filter (Graph 400); a $search
      // error here means no usable result, so return empty rather than a mis-broadened $filter.
      if (usesKqlOnly || isKQLFormat(query || '')) {
        return { isError: !error.isEmptyResult, content: [{ type: "text", text:
          `Search returned no usable results (${degradedBecause}; KQL or from/to/hasAttachments/`
          + `importance predicates cannot fall back to a subject filter, so no second attempt was made).`
          + (error.isEmptyResult ? '' : ' The mailbox search failed; this is not evidence of absence.') }] };
      }

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
        maxResults,
        degradedBecause
      });
    }
  } catch (error) {
    console.error('Error in unified email search:', error);
    return {
      isError: true,
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
    const froms = Array.isArray(params.from) ? params.from : [params.from];
    kql.push(froms.length === 1 ? `from:${froms[0]}` : `(${froms.map((f) => `from:${f}`).join(' OR ')})`);
  }
  if (params.to) {
    const tos = Array.isArray(params.to) ? params.to : [params.to];
    kql.push(tos.length === 1 ? `to:${tos[0]}` : `(${tos.map((t) => `to:${t}`).join(' OR ')})`);
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
function isKQLFormat(query = '') {
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
    // Fall through to the next tier rather than terminating. Tier 1 only runs for
    // "complex" (e.g. multi-word) queries, so returning a terminal empty result here
    // meant multi-word searches never reached the $search or $filter tiers at all,
    // while single-word searches — which skip Tier 1 — found matches fine. The caller
    // saw a confident "No emails found" for a mailbox that did contain the message.
    // Tier 2 does the same thing, and the caller already catches and falls back.
    console.error('Microsoft Search API returned 0 hits, falling back');
    throw new Error('No results from Microsoft Search API, trigger fallback');
  }

  const emails = hits.map(hit => summarizeEmail(hit.resource, hit.hitId));
  const emailsList = emails.map((email, index) => formatSearchEmail(email, hits[index].resource)).join('\n');

  const sortNote = useRelevance ? ' (sorted by relevance)' : ' (sorted by date)';
  const container = response.value[0]?.hitsContainers[0];
  const coverage = searchCoverage('microsoft-search', hits.length, Math.min(maxResults, 1000),
    !!container?.moreResultsAvailable || container?.total > hits.length);

  return withMailStructuredContent({
    structuredContent: { ...coverage, emails },
    content: [{
      type: "text",
      text: `Found ${hits.length} emails${sortNote}:\n${coverage.note}\n\n${emailsList}`
    }]
  });
}

/**
 * Search using Graph API $search parameter
 */
async function searchUsingGraphSearch(accessToken, params) {
  const { query, endpoint, maxResults, allowEmpty = false, mailbox } = params;

  const queryParams = {
    // Quote the OData string AND escape its inner KQL phrase quotes. Merely
    // wrapping a query such as subject:"Banthar" produces an invalid request.
    $search: JSON.stringify(query),
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
    if (allowEmpty) {
      // Predicates like from:/to:/hasattachment:/importance: are only expressible via $search
      // (a $filter on them triggers a Graph "too complex" 400), so there is no $filter fallback
      // to attempt — a genuine empty result is the answer.
      return {
        content: [{ type: "text", text: "No emails found matching your search (bodies were searched)." }]
      };
    }
    console.error('Graph $search returned 0 results, falling back to $filter');
    // Flagged so the caller can tell "Graph looked and found nothing" apart from "Graph
    // errored and we never looked". Both used to arrive as the same bare throw.
    const empty = new Error('No results from $search, trigger fallback');
    empty.isEmptyResult = true;
    throw empty;
  }

  const emails = response.value.map(email => summarizeEmail(email));
  const emailsList = emails.map((email, index) => formatSearchEmail(email, response.value[index])).join('\n');

  const coverage = searchCoverage('graph-search', response.value.length, queryParams.$top,
    !!response['@odata.nextLink']);
  return withMailStructuredContent({
    structuredContent: { ...coverage, emails },
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails:\n${coverage.note}\n\n${emailsList}`
    }]
  });
}

/**
 * Search using $filter parameter (most reliable fallback)
 */
async function searchUsingFilter(accessToken, params) {
  // NOTE: from/to/hasAttachments/importance are intentionally NOT handled here. They are routed
  // through the KQL $search path upstream (handleEmailSearch), because a Graph message $filter on
  // those properties returns "The restriction or sort order is too complex for this operation"
  // (and to/from fail the filter node outright). This function builds $filter only for text
  // (contains), isRead, and date-range predicates.
  const {
    query,
    subject,
    isRead,
    startDate,
    endDate,
    endpoint,
    maxResults
  , mailbox, degradedBecause } = params;

  let filters = [];

  // Tokenise multi-word queries. `contains()` is a literal substring match, so a
  // multi-word query used whole ("Air Products PRISM membrane") only matches a subject
  // containing that exact run of characters — which almost never happens, and the caller
  // gets a bare "No emails found" that reads like an empty mailbox. AND the terms
  // instead so word order and intervening words don't matter.
  const termFilters = (text) =>
    String(text)
      .split(/\s+/)
      .filter(Boolean)
      .map(t => `contains(subject, '${t.replace(/'/g, "''")}')`);

  if (query && query !== '' && query !== '*' && !subject) {
    filters.push(...termFilters(query));
  }

  if (subject) {
    filters.push(...termFilters(subject));
  }

  if (isRead !== undefined) {
    filters.push(`isRead eq ${isRead ? 'true' : 'false'}`);
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
    // Say what was actually searched. This tier matches on SUBJECT only, so a miss here
    // does not mean the mailbox has nothing — it means no subject line matched. Without
    // this, callers read a bare "No emails found" as an authoritative empty result.
    const searched = filterQuery
      ? `subject-line match on: ${filterQuery}`
      : 'no text predicate';
    const why = degradedBecause ? `\nFell back because ${degradedBecause}.` : '';
    return {
      content: [{
        type: "text",
        text: `No emails found (filter fallback searched SUBJECT LINES ONLY — ${searched}).${why}\n`
            + `Message bodies were NOT searched, so this is not evidence the mailbox has nothing. `
            + `Try a single distinctive keyword, or the 'from' parameter to search by sender.`
      }]
    };
  }

  const emails = response.value.map(email => summarizeEmail(email));
  const emailsList = emails.map((email, index) => formatSearchEmail(email, response.value[index])).join('\n');

  const why = degradedBecause ? ` — fell back because ${degradedBecause}` : '';
  const coverage = searchCoverage('subject-filter', response.value.length, queryParams.$top,
    !!response['@odata.nextLink']);
  return withMailStructuredContent({
    structuredContent: { ...coverage, emails },
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails (filter fallback, SUBJECT LINES ONLY${why}).\n`
          + `Message bodies were NOT searched, so matches that appear only in a body are missing `
          + `from this list.\n${coverage.note}\n\n${emailsList}`
    }]
  });
}

function summarizeEmail(email, fallbackId) {
  const from = email.from?.emailAddress?.address
    || (typeof email.from === 'string' ? email.from : null)
    || 'Unknown sender';
  return {
    id: email.id || fallbackId || 'Not available',
    subject: email.subject || '(No subject)',
    from,
    receivedDateTime: email.receivedDateTime || null,
    hasAttachments: !!email.hasAttachments
  };
}

function formatSearchEmail(email, source) {
  const attachments = email.hasAttachments ? ' 📎' : '';
  const importance = source.importance ? ` [${source.importance}]` : '';
  const unread = source.isRead === false ? ' *' : '';
  const date = email.receivedDateTime ? new Date(email.receivedDateTime).toLocaleString() : 'Unknown date';
  return `- ${email.subject}${attachments}${importance}${unread}\n  From: ${email.from}\n  Date: ${date}\n  ID: ${email.id}\n`;
}

// A result page is not a completeness claim. Graph message $search also has
// a service ceiling of 1,000 results even when the mailbox has more matches.
function searchCoverage(method, returned, pageLimit, hasMore) {
  const limited = hasMore || returned >= pageLimit;
  return {
    method, returned, pageLimit, hasMore, limited,
    serviceLimit: method === 'graph-search' ? 1000 : null,
    note: limited
      ? `PARTIAL SEARCH: returned one page (limit ${pageLimit}); additional matches may exist. Narrow the date range or query before claiming completeness.`
      : 'One result page returned; search-index results are not a mailbox completeness guarantee.'
  };
}

// Deprecated search functions kept for internal compatibility
async function searchEmailsBasic(accessToken, params) {
  const { query, from, subject, maxResults, mailbox } = params;

  let searchQuery = query;
  if (from) searchQuery = `from:${from} AND ${searchQuery}`;
  if (subject) searchQuery = `subject:${subject} AND ${searchQuery}`;

  const queryParams = {
    $search: JSON.stringify(searchQuery),
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

  const emails = response.value.map(email => summarizeEmail(email));
  const emailsList = emails.map((email, index) => formatSearchEmail(email, response.value[index])).join('\n');

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

  const emails = response.value.map(email => summarizeEmail(email));
  const emailsList = emails.map((email, index) => formatSearchEmail(email, response.value[index])).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails:\n\n${emailsList}`
    }]
  };
}

module.exports = { handleEmailSearch };
