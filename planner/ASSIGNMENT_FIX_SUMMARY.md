# Planner Task Assignment Fix Summary

## Issues Fixed

1. **Custom Headers Not Handled in callGraphAPI**
   - Updated `callGraphAPI` function to accept a `customHeaders` parameter
   - Headers are now properly merged into the request options

2. **Assignment Format Issues**
   - For task creation: Uses `orderHint` without `@odata.type`
   - For task updates: Uses both `orderHint` and `@odata.type`
   - Properly handles If-Match headers for PATCH and DELETE operations

3. **Error Handling Improvements**
   - Enhanced error messages to parse and display actual API error responses
   - Better error propagation with original error messages preserved

## Files Modified

### 1. /mnt/c/Users/hvksh/mcp-servers/office-mcp/utils/graph-api.js
- Added `customHeaders` parameter to callGraphAPI function
- Updated headers merging logic
- Improved error response parsing

### 2. /mnt/c/Users/hvksh/mcp-servers/office-mcp/planner/tasks.js
- Fixed assignment format for create operations (orderHint only)
- Fixed assignment format for update operations (orderHint + @odata.type)
- Added missing null parameter for queryParams in PATCH/DELETE calls
- Enhanced error messages for better debugging

## New Files Created

### 1. /mnt/c/Users/hvksh/mcp-servers/office-mcp/planner/tasks-enhanced.js
- Enhanced task management functions with better error handling
- Separate functions for creating tasks with assignments and updating assignments
- Better assignment management with add/remove functionality

### 2. /mnt/c/Users/hvksh/mcp-servers/office-mcp/test-assignment-fix.js
- Test file to verify the assignment functionality works correctly
- Tests for single and multiple assignments
- Tests for updating existing task assignments

## Key Changes

1. **callGraphAPI signature changed**:
   ```javascript
   // Before
   async function callGraphAPI(accessToken, method, path, data = null, queryParams = {})
   
   // After
   async function callGraphAPI(accessToken, method, path, data = null, queryParams = {}, customHeaders = {})
   ```

2. **Assignment format for creation**:
   ```javascript
   taskData.assignments[userId] = {
     orderHint: ' !'
   };
   ```

3. **Assignment format for updates**:
   ```javascript
   updateData.assignments[userId] = { 
     orderHint: ' !',
     '@odata.type': '#microsoft.graph.plannerAssignment' 
   };
   ```

4. **Header passing for PATCH/DELETE**:
   ```javascript
   await callGraphAPI(
     accessToken,
     'PATCH',
     `planner/tasks/${taskId}`,
     updateData,
     null, // queryParams
     {
       'If-Match': currentTask['@odata.etag']
     }
   );
   ```

## Testing

To test the fixes, use the test file:
```bash
node test-assignment-fix.js
```

Or use the enhanced tools directly in your main application by importing:
```javascript
const { enhancedTaskTools } = require('./planner/tasks-enhanced');
```

The fixes ensure proper task assignment handling in Microsoft Planner through the Graph API.