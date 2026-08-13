/**
 * Microsoft Graph API helper functions with enhanced error handling
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

// Retry configuration
const RETRY_CONFIG = {
  maxRetries: 3,
  retryDelay: 1000, // Start with 1 second
  retryableErrors: [429, 503, 504], // Rate limit and service unavailable
  exponentialBackoff: true
};

const MAX_REDIRECTS = 5;

// Error message enhancements
const ERROR_SUGGESTIONS = {
  401: 'Authentication token may have expired. Please re-authenticate.',
  403: 'Insufficient permissions. Check if the app has the required Microsoft Graph permissions.',
  404: 'Resource not found. Verify the ID or path is correct.',
  429: 'Rate limit exceeded. The request will be retried automatically.',
  500: 'Microsoft Graph service error. Please try again later.',
  503: 'Service temporarily unavailable. The request will be retried automatically.'
};

const MAILBOX_SCOPED_RESOURCES = new Set([
  'messages',
  'mailFolders',
  'sendMail',
  'mailboxSettings'
]);

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getHeaderValue(headers, name) {
  const match = Object.keys(headers || {}).find(key => key.toLowerCase() === name.toLowerCase());
  return match ? headers[match] : undefined;
}

function hasHeader(headers, name) {
  return getHeaderValue(headers, name) !== undefined;
}

function getAnchorMailboxFromPath(path) {
  const match = String(path || '').match(/^\/?users\/([^/?#]+)(?:\/([^/?#]+))?/i);
  if (!match || !match[2] || !MAILBOX_SCOPED_RESOURCES.has(match[2])) {
    return null;
  }

  try {
    return decodeURIComponent(match[1]);
  } catch (error) {
    return match[1];
  }
}

function normalizeGraphError(statusCode, parsedError, rawBody, responseHeaders = {}) {
  const graphBody = parsedError?.error || parsedError || {};
  const innerError = graphBody.innerError || graphBody.innererror || null;
  const requestId = innerError?.['request-id'] ||
    innerError?.requestId ||
    responseHeaders['request-id'] ||
    responseHeaders['client-request-id'] ||
    null;

  return {
    statusCode,
    code: graphBody.code || null,
    message: graphBody.message || (typeof rawBody === 'string' ? rawBody : null),
    innerError,
    requestId,
    date: innerError?.date || responseHeaders.date || null
  };
}

function createGraphAPIError(statusCode, responseData, parsedError, responseHeaders = {}) {
  const graphError = normalizeGraphError(statusCode, parsedError, responseData, responseHeaders);
  const errorMessage = graphError.message || responseData;
  const suggestion = ERROR_SUGGESTIONS[statusCode] || '';
  const fullMessage = suggestion
    ? `API call failed with status ${statusCode}: ${errorMessage}\nSuggestion: ${suggestion}`
    : `API call failed with status ${statusCode}: ${errorMessage}`;

  const error = new Error(fullMessage);
  error.statusCode = statusCode;
  error.graphError = graphError;
  return error;
}

/**
 * Encode Graph API path segments intelligently
 * Preserves OData operators ($value, $ref, $count) and drive path syntax (:/path:/content)
 * @param {string} path - The API path to encode
 * @returns {string} - Properly encoded path
 */
function encodeGraphPath(path) {
  if (!path) return path;

  // If path already contains encoded characters, assume it's already encoded
  if (path.includes('%')) return path;

  // Split path and process each segment
  return path.split('/').map(segment => {
    // Don't encode empty segments
    if (!segment) return segment;

    // Preserve OData operators: $value, $ref, $count, $batch, etc.
    if (/^\$[a-zA-Z]+$/.test(segment)) return segment;

    // Preserve drive path syntax that starts with colon (e.g., "root:", ":path", ":/Documents")
    // These include OneDrive/SharePoint path-based addressing
    if (segment.startsWith(':') || segment.endsWith(':')) return segment;

    // Preserve segments that are purely alphanumeric with hyphens/underscores (IDs, simple names)
    if (/^[a-zA-Z0-9_-]+$/.test(segment)) return segment;

    // Preserve Graph function calls like search(q='...')
    if (segment.includes('(') && segment.includes(')')) return segment;

    // For other segments, encode but preserve colons for drive path syntax
    return encodeURIComponent(segment).replace(/%3A/g, ':');
  }).join('/');
}

/**
 * Makes a request to the Microsoft Graph API with retry logic
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {string} path - API endpoint path
 * @param {object} data - Data to send for POST/PUT requests
 * @param {object} queryParams - Query parameters
 * @param {object} customHeaders - Custom headers
 * @param {number} retryCount - Current retry attempt (internal use)
 * @returns {Promise<object>} - The API response
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, customHeaders = {}, retryCount = 0, redirectCount = 0) {
  // For test tokens, we'll simulate the API call
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`[GRAPH-API] TEST MODE: ${method} ${path}`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  try {
    // Safe logging - only log path without sensitive query params
    const isAbsoluteUrl = /^https?:\/\//i.test(path);
    console.error(`[GRAPH-API] ${method} ${isAbsoluteUrl ? '[external-url]' : path.split('?')[0]}`);

    // Clone queryParams to avoid mutating caller's object
    const params = { ...queryParams };

    // Encode path using Graph-aware encoder
    const encodedPath = isAbsoluteUrl ? path : encodeGraphPath(path);
    
    // Build query string from parameters with special handling for OData filters
    let queryString = '';
    if (params && Object.keys(params).length > 0) {
      // Handle $filter parameter specially to ensure proper URI encoding
      const filter = params.$filter;
      if (filter) {
        delete params.$filter; // Remove from regular params
      }

      // Build query string with proper encoding for regular params
      const urlParams = new URLSearchParams();
      for (const [key, value] of Object.entries(params)) {
        urlParams.append(key, value);
      }

      queryString = urlParams.toString();

      // Add filter parameter separately with proper encoding
      if (filter) {
        if (queryString) {
          queryString += `&$filter=${encodeURIComponent(filter)}`;
        } else {
          queryString = `$filter=${encodeURIComponent(filter)}`;
        }
      }

      if (queryString) {
        queryString = '?' + queryString;
      }

      // Only log query param count, not content (may contain sensitive data)
      if (process.env.DEBUG_VERBOSE === 'true') {
        console.error(`[GRAPH-API] Query params: ${Object.keys(params).length} params`);
      }
    }

    const url = isAbsoluteUrl ? `${encodedPath}${queryString}` : `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
    // Only log full URL in verbose debug mode (may contain sensitive data)
    if (process.env.DEBUG_VERBOSE === 'true') {
      console.error(`[GRAPH-API] Full URL: ${url}`);
    }
    
    customHeaders = { ...(customHeaders || {}) };

    const anchorMailbox = getAnchorMailboxFromPath(path);
    if (anchorMailbox && !hasHeader(customHeaders, 'X-AnchorMailbox')) {
      customHeaders['X-AnchorMailbox'] = anchorMailbox;
    }

    // Determine if this is a binary request/response based on Content-Type
    const requestContentType = getHeaderValue(customHeaders, 'Content-Type') || 'application/json';
    const isBinaryRequest = requestContentType.includes('application/octet-stream') ||
                           requestContentType.includes('image/') ||
                           requestContentType.includes('audio/') ||
                           requestContentType.includes('video/') ||
                           Buffer.isBuffer(data);

    // Inject Prefer: outlook.timezone header on calendar GET requests
    // so Graph returns event times in the configured timezone
    if (method === 'GET' && !customHeaders['Prefer']) {
      const pathLower = path.toLowerCase();
      if (pathLower.includes('/calendar') || pathLower.includes('/events') || pathLower.includes('/calendarview')) {
        customHeaders = { ...customHeaders, 'Prefer': `outlook.timezone="${config.getMsTimezone()}"` };
      }
    }

    return new Promise((resolve, reject) => {
      const requestHost = new URL(url).hostname;
      const graphHost = new URL(config.GRAPH_API_ENDPOINT).hostname;
      const shouldSendAuth = requestHost === graphHost;
      const options = {
        method: method,
        headers: {
          'Content-Type': isBinaryRequest ? requestContentType : 'application/json',
          ...(shouldSendAuth ? { 'Authorization': `Bearer ${accessToken}` } : {}),
          ...customHeaders
        }
      };

      const req = https.request(url, options, (res) => {
        const responseContentType = res.headers['content-type'] || '';
        const location = res.headers.location;

        if ([301, 302, 303, 307, 308].includes(res.statusCode) && location) {
          res.resume();

          if (!['GET', 'HEAD'].includes(method)) {
            reject(new Error(`Redirect received for unsupported method ${method}`));
            return;
          }

          if (redirectCount >= MAX_REDIRECTS) {
            reject(new Error(`Too many redirects while calling Microsoft Graph (${MAX_REDIRECTS})`));
            return;
          }

          const redirectUrl = new URL(location, url).toString();
          console.error(`[GRAPH-API] Following redirect ${redirectCount + 1}/${MAX_REDIRECTS}`);
          callGraphAPI(accessToken, method, redirectUrl, data, {}, customHeaders, retryCount, redirectCount + 1)
            .then(resolve)
            .catch(reject);
          return;
        }

        // Determine if the response should be handled as binary.
        //
        // FAIL-CLOSED: default to binary unless the content type is provably text.
        // The previous allow-list (octet-stream / image / audio / video / pdf /
        // vnd.openxmlformats) missed the legacy and macro-enabled Office types --
        // application/vnd.ms-excel.sheet.macroEnabled.12 (.xlsm), application/msword,
        // application/vnd.ms-powerpoint, application/zip, ... -- which fell through to
        // Buffer.concat(chunks).toString('utf8'). Decoding binary as UTF-8 replaces
        // every invalid byte with U+FFFD, which is IRREVERSIBLE: a 3.6 MB .xlsm
        // arrived as 6.8 MB with 1.55M mangled bytes and an unreadable zip directory.
        // Misclassifying text as binary, by contrast, is harmless -- the caller gets
        // a Buffer saved to a temp file instead of inline text, losing nothing.
        // JSON parsing below is unaffected (json is classified as text here).
        const isTextResponse = responseContentType.includes('json') ||
                               responseContentType.startsWith('text/') ||
                               responseContentType.includes('xml') ||
                               responseContentType.includes('csv') ||
                               responseContentType.includes('javascript') ||
                               responseContentType.includes('x-www-form-urlencoded');
        const isBinaryResponse = !isTextResponse;

        // Collect response data - use array of buffers for binary, string for text
        const chunks = [];

        res.on('data', (chunk) => {
          chunks.push(chunk);
        });

        res.on('end', () => {
          // Combine chunks appropriately. Empty bodies (204s etc.) stay on the string
          // path regardless of classification: callers test `!response`, and an empty
          // Buffer is truthy where the empty string they used to get is falsy.
          const rawBody = Buffer.concat(chunks);
          const responseData = (isBinaryResponse && rawBody.length > 0)
            ? rawBody
            : rawBody.toString('utf8');

          if (res.statusCode >= 200 && res.statusCode < 300) {
            // Handle binary responses
            if (isBinaryResponse) {
              resolve(responseData); // Return Buffer
            }
            // Handle text/non-JSON responses (like WEBVTT transcripts)
            else if (responseContentType.includes('text/vtt') ||
                     responseContentType.includes('text/plain') ||
                     !responseContentType.includes('json')) {
              resolve(responseData); // Return string
            }
            // Parse JSON responses
            else {
              try {
                const jsonResponse = JSON.parse(responseData);
                resolve(jsonResponse);
              } catch (error) {
                reject(new Error(`Error parsing API response: ${error.message}`));
              }
            }
          } else {
            // Handle other errors with retry logic
            const shouldRetry = RETRY_CONFIG.retryableErrors.includes(res.statusCode) && 
                              retryCount < RETRY_CONFIG.maxRetries;
            
            if (shouldRetry) {
              // Calculate delay with exponential backoff
              const delay = RETRY_CONFIG.exponentialBackoff 
                ? RETRY_CONFIG.retryDelay * Math.pow(2, retryCount)
                : RETRY_CONFIG.retryDelay;
              
              console.error(`Request failed with status ${res.statusCode}. Retrying in ${delay}ms... (Attempt ${retryCount + 1}/${RETRY_CONFIG.maxRetries})`);
              
              // Wait and retry
              setTimeout(() => {
                callGraphAPI(accessToken, method, path, data, queryParams, customHeaders, retryCount + 1)
                  .then(resolve)
                  .catch(reject);
              }, delay);
              return;
            }
            
            // Parse error and add suggestions while preserving the structured Graph error.
            try {
              const errorData = JSON.parse(responseData);
              reject(createGraphAPIError(res.statusCode, responseData, errorData, res.headers));
            } catch (parseError) {
              reject(createGraphAPIError(res.statusCode, responseData, null, res.headers));
            }
          }
        });
      });
      
      req.on('error', (error) => {
        reject(new Error(`Network error during API call: ${error.message}`));
      });

      // Write request body if present
      if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
        if (isBinaryRequest || Buffer.isBuffer(data)) {
          // Write binary data directly (Buffer or raw bytes)
          req.write(Buffer.isBuffer(data) ? data : Buffer.from(data));
        } else {
          // Serialize JSON data
          req.write(JSON.stringify(data));
        }
      }

      req.end();
    });
  } catch (error) {
    console.error('[GRAPH-API] Error:', error.message);
    throw error;
  }
}

module.exports = {
  callGraphAPI,
  getAnchorMailboxFromPath,
  createGraphAPIError
};
