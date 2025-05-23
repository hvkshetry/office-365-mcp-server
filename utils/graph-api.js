/**
 * Microsoft Graph API helper functions
 */
const https = require('https');
const config = require('../config');
const mockData = require('./mock-data');

/**
 * Makes a request to the Microsoft Graph API
 * @param {string} accessToken - The access token for authentication
 * @param {string} method - HTTP method (GET, POST, etc.)
 * @param {string} path - API endpoint path
 * @param {object} data - Data to send for POST/PUT requests
 * @param {object} queryParams - Query parameters
 * @returns {Promise<object>} - The API response
 */
async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, customHeaders = {}) {
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
            reject(new Error('UNAUTHORIZED'));
          } else {
            try {
              const errorData = JSON.parse(responseData);
              const errorMessage = errorData.error?.message || responseData;
              reject(new Error(`API call failed with status ${res.statusCode}: ${errorMessage}`));
            } catch (parseError) {
              reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
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
