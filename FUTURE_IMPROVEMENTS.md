# Future Improvements for Office MCP

This document tracks potential improvements and feature requests for the Office MCP (Microsoft Cloud Platform) integration.

## Calendar Management

### 1. Recurring Meeting Support
**Priority:** High  
**Date Identified:** May 23, 2025  
**Requested By:** User feedback during bookkeeping meeting setup  

**Current Limitation:**
- The `calendar` tool currently only supports creating individual calendar events
- No built-in parameter for recurrence patterns (daily, weekly, monthly, quarterly, etc.)
- Workaround requires creating multiple individual events manually

**Desired Functionality:**
- Add recurrence parameters to the `calendar` create operation:
  ```javascript
  calendar {
    operation: "create",
    subject: "Recurring Meeting",
    start: "2025-07-02T18:00:00Z",
    end: "2025-07-02T19:00:00Z",
    recurrence: {
      pattern: "quarterly", // or "daily", "weekly", "monthly", "yearly"
      interval: 1,
      count: 4, // or endDate: "2026-12-31"
      daysOfWeek: ["wednesday"], // for weekly patterns
      dayOfMonth: 1, // for monthly patterns
    }
  }
  ```

**Benefits:**
- Reduces manual work for scheduling regular meetings
- Ensures consistency in meeting schedules
- Matches user expectations from Outlook/Teams UI
- Prevents scheduling errors and missed meetings

**Implementation Notes:**
- Microsoft Graph API supports recurrence through the `recurrence` property
- Would need to map simplified parameters to Graph API recurrence pattern format
- Should support common patterns: daily, weekly, monthly, quarterly, yearly

---

## Email Management

### 2. ✅ COMPLETED: Unified Email Search with Advanced Capabilities
**Priority:** High  
**Date Completed:** August 24, 2025  
**Implemented By:** Assistant with user hvksh

**Previous Limitations:**
- Three separate search modes (basic, enhanced, simple) causing confusion
- Limited KQL (Keyword Query Language) support
- No folder-specific search capability
- No date range filtering
- Limited filter combinations
- No relevance ranking option

**Implemented Solution:**
- Single unified `email_search` tool with automatic optimization
- Full KQL syntax support with auto-detection
- Folder search by name or ID with well-known folder mapping
- Date range filtering with relative dates (7d, 1w, 1m, 1y)
- Comprehensive filter parameters (hasAttachments, isRead, importance, etc.)
- Three-tier execution strategy: Microsoft Search API → $search → $filter
- Relevance ranking option using Microsoft Search API
- Smart query builder that converts natural language to KQL

**Key Features Added:**
1. **KQL Support**: Full support for complex queries like `from:john@example.com AND (subject:report OR body:analysis)`
2. **Folder Search**: Search specific folders by name (`folderName: "inbox"`) or ID
3. **Date Filtering**: Support for both ISO dates and relative dates (`startDate: "7d"`)
4. **Smart Fallbacks**: Automatic fallback from advanced to simple search methods
5. **Relevance Ranking**: Option to sort by relevance instead of date
6. **Natural Language**: Automatically converts simple queries to KQL format

**Files Modified:**
- `/mnt/c/Users/hvksh/mcp-servers/office-mcp/email/index.js` - Complete rewrite of search functionality
- `/home/hvksh/admin/CLAUDE.md` - Updated with AI-optimized documentation

---

### 3. Batch Email Operations
**Priority:** Medium  
**Date Identified:** May 23, 2025  

**Current Limitation:**
- Email operations are primarily single-email focused
- Batch operations require multiple API calls

**Desired Functionality:**
- Support for batch email operations (mark as read, move multiple, delete multiple)
- Better performance for bulk email management

---

## Teams Integration

### 3. Meeting Recording Management
**Priority:** Medium  
**Date Identified:** May 23, 2025  

**Current Limitation:**
- Can retrieve recordings but limited management capabilities
- No ability to download or share recordings programmatically

**Desired Functionality:**
- Download meeting recordings
- Share recordings with specific users
- Set expiration dates on recordings

---

## Planner/Tasks

### 4. Task Dependencies
**Priority:** Low  
**Date Identified:** May 23, 2025  

**Current Limitation:**
- No support for task dependencies in Planner
- Cannot link related tasks

**Desired Functionality:**
- Create task dependencies
- Visualize task relationships
- Automatic date adjustments based on dependencies

---

## General Improvements

### 5. Error Message Enhancement
**Priority:** Medium  
**Date Identified:** May 23, 2025  

**Current Limitation:**
- Some error messages are technical and not user-friendly
- Difficult to understand root cause of failures

**Desired Functionality:**
- More descriptive error messages
- Suggested actions for common errors
- Better handling of authentication timeouts

---

## How to Contribute

If you identify additional improvements or limitations:

1. Add them to this document with:
   - Clear description of current limitation
   - Desired functionality
   - Use case/benefit
   - Priority (High/Medium/Low)
   - Date identified

2. Follow the format above for consistency

3. Consider implementation complexity and API limitations

---

Last Updated: August 24, 2025