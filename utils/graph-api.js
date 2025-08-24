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

// Error message enhancements
const ERROR_SUGGESTIONS = {
  401: 'Authentication token may have expired. Please re-authenticate.',
  403: 'Insufficient permissions. Check if the app has the required Microsoft Graph permissions.',
  404: 'Resource not found. Verify the ID or path is correct.',
  429: 'Rate limit exceeded. The request will be retried automatically.',
  500: 'Microsoft Graph service error. Please try again later.',
  503: 'Service temporarily unavailable. The request will be retried automatically.'
};

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
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
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, customHeaders = {}, retryCount = 0) {
  // For test tokens, we'll simulate the API call
  if (config.USE_TEST_MODE && accessToken.startsWith('test_access_token_')) {
    console.error(`TEST MODE: Simulating ${method} ${path} API call`);
    return mockData.simulateGraphAPIResponse(method, path, data, queryParams);
  }

  try {
    console.error(`Making real API call: ${method} ${path}`);
    
    // Encode path segments properly
    const encodedPath = path.split('/')
      .map(segment => encodeURIComponent(segment))
      .join('/');
    
    // Build query string from parameters with special handling for OData filters
    let queryString = '';
    if (queryParams && Object.keys(queryParams).length > 0) {
      // Handle $filter parameter specially to ensure proper URI encoding
      const filter = queryParams.$filter;
      if (filter) {
        delete queryParams.$filter; // Remove from regular params
      }
      
      // Build query string with proper encoding for regular params
      const params = new URLSearchParams();
      for (const [key, value] of Object.entries(queryParams)) {
        params.append(key, value);
      }
      
      queryString = params.toString();
      
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
      
      console.error(`Query string: ${queryString}`);
    }
    
    const url = `${config.GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
    console.error(`Full URL: ${url}`);
    
    return new Promise((resolve, reject) => {
      const options = {
        method: method,
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
          ...customHeaders
        }
      };
      
      const req = https.request(url, options, (res) => {
        let responseData = '';
        
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        
        res.on('end', () => {
          if (res.statusCode >= 200 && res.statusCode < 300) {
            const contentType = res.headers['content-type'] || '';
            
            // Handle non-JSON responses (like WEBVTT transcripts)
            if (contentType.includes('text/vtt') || contentType.includes('text/plain') || 
                contentType.includes('application/vnd.openxmlformats') || 
                !contentType.includes('json')) {
              // Return raw text for transcript content and other non-JSON responses
              resolve(responseData);
            } else {
              // Parse JSON responses
              try {
                const jsonResponse = JSON.parse(responseData);
                resolve(jsonResponse);
              } catch (error) {
                reject(new Error(`Error parsing API response: ${error.message}`));
              }
            }
          } else if (res.statusCode === 401) {
            // Token expired or invalid
            const suggestion = ERROR_SUGGESTIONS[401];
            reject(new Error(`UNAUTHORIZED: ${suggestion}`));
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
            
            // Parse error and add suggestions
            try {
              const errorData = JSON.parse(responseData);
              const errorMessage = errorData.error?.message || responseData;
              const suggestion = ERROR_SUGGESTIONS[res.statusCode] || '';
              const fullMessage = suggestion 
                ? `API call failed with status ${res.statusCode}: ${errorMessage}\nSuggestion: ${suggestion}`
                : `API call failed with status ${res.statusCode}: ${errorMessage}`;
              reject(new Error(fullMessage));
            } catch (parseError) {
              const suggestion = ERROR_SUGGESTIONS[res.statusCode] || '';
              const fullMessage = suggestion
                ? `API call failed with status ${res.statusCode}: ${responseData}\nSuggestion: ${suggestion}`
                : `API call failed with status ${res.statusCode}: ${responseData}`;
              reject(new Error(fullMessage));
            }
          }
        });
      });
      
      req.on('error', (error) => {
        reject(new Error(`Network error during API call: ${error.message}`));
      });
      
      if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
        req.write(JSON.stringify(data));
      }
      
      req.end();
    });
  } catch (error) {
    console.error('Error calling Graph API:', error);
    throw error;
  }
}

module.exports = {
  callGraphAPI
};
