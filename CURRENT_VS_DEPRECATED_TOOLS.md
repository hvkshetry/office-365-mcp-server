# Office MCP Server - Current vs. Deprecated Tools Documentation

## Overview

The Office MCP server has recently undergone a significant tool consolidation to improve usability while maintaining full functionality. This document details the transition from the previous approach with many individual tools to the current consolidated approach with operation-based tools.

## Tool Count Reduction

- **Original Tool Count**: ~100+ tools
- **Current Tool Count**: 24 tools
- **Reduction**: ~76% fewer tools

## Consolidated Tools Architecture

The consolidated tools follow a consistent pattern:

```json
{
  "operation": "list|create|update|delete|etc",
  // ... other parameters specific to the operation
}
```

This approach:
- Reduces cognitive load by providing fewer tools
- Maintains discoverability by grouping related operations in a single tool
- Preserves all original functionality through the operation parameter

## Current (Consolidated) vs. Deprecated (Individual) Tools

### Authentication Module

**Status**: Unchanged (3 tools)

The Authentication module tools remain the same:
- `authenticate`
- `check_auth_status`
- `logout`

### Email Module

**Status**: Consolidated (from 11 tools to 5 tools)

| Current (Consolidated) Tool | Deprecated (Individual) Tools | Migration Notes |
|----------------------------|------------------------------|----------------|
| `email` <br> (with operations: list, read, send) | `list_emails` <br> `read_email` <br> `send_email` | Use the `operation` parameter to specify action |
| `email_search` <br> (with modes: basic, enhanced, simple) | `search_emails` <br> `search_emails_enhanced` <br> `search_emails_simple` | Use the `mode` parameter to specify search type |
| `email_move` | `move_email` <br> `batch_move_emails` | Automatic batch mode for >5 emails |
| `email_folder` <br> (with operations: list, create) | `list_email_folders` <br> `create_email_folder` | Use the `operation` parameter |
| `email_rules` <br> (with operations: list, create) | `list_email_rules` <br> `create_email_rule` | Supports both standard and enhanced modes |

### Calendar Module

**Status**: Consolidated (from 5 tools to 1 tool)

| Current (Consolidated) Tool | Deprecated (Individual) Tools | Migration Notes |
|----------------------------|------------------------------|----------------|
| `calendar` <br> (with operations: list, create, get, update, delete) | `calendar_list_events` <br> `calendar_create_event` <br> `calendar_get_event` <br> `calendar_update_event` <br> `calendar_delete_event` | Use the `operation` parameter to specify the calendar action |

### Teams Module

**Status**: Fully Consolidated (from 25+ tools to 3 tools)

| Current (Consolidated) Tool | Removed (Individual) Tools | Migration Notes |
|----------------------------|------------------------------|----------------|
| `teams_meeting` | `teams`, `teams_create_meeting`, `teams_update_meeting`, `teams_cancel_meeting`, `teams_get_meeting`, `teams_find_meeting_by_url`, `teams_list_transcripts`, `teams_get_transcript`, `teams_list_recordings`, `teams_get_recording`, `teams_get_participants`, `teams_get_insights` | Use `operation` parameter: "create", "update", "cancel", "get", "find_by_url", "list_transcripts", "get_transcript", "list_recordings", "get_recording", "get_participants", "get_insights" |
| `teams_channel` | `channels`, `channel_messages`, `channel_members`, `tabs`, and all individual channel operations | Use `operation` parameter: "list", "create", "get", "update", "delete", "list_messages", "get_message", "create_message", "reply_to_message", "list_members", "add_member", "remove_member", "list_tabs" |
| `teams_chat` | `chat`, `chat_members`, and all individual chat operations | Use `operation` parameter: "list", "create", "get", "update", "delete", "list_messages", "get_message", "send_message", "update_message", "delete_message", "list_members", "add_member", "remove_member" |

**Legacy Tools Removed**: `teams`, `channels`, `channel_messages`, `channel_members`, `chat`, `chat_members`, `tabs`, `activity_feed`, `team_archive`

### Planner Module

**Status**: Consolidated (from 18 tools to 8 tools)

| Current (Consolidated) Tool | Deprecated (Individual) Tools | Migration Notes |
|----------------------------|------------------------------|----------------|
| `planner_plan` | `planner_list_plans` <br> `planner_get_plan` <br> `planner_create_plan` <br> `planner_update_plan` <br> `planner_delete_plan` | Use appropriate `operation` parameter |
| `planner_task` | `planner_list_tasks` <br> `planner_get_task` <br> `planner_create_task` <br> `planner_update_task` <br> `planner_delete_task` | Use appropriate `operation` parameter |
| Additional planner tools | Various specialized planner tools | Several specialized tools remain for complex operations |

### Notifications Module

**Status**: Unchanged (4 tools)

The Notifications module tools remain the same:
- `notification_create_subscription`
- `notification_list_subscriptions`
- `notification_get_subscription`
- `notification_delete_subscription`

### Removed Modules

- **Drive Module**: Completely removed (not needed given local file access)
- **Users Module**: Completely removed (not needed for executive assistant role)

## Common Issues During Transition

Some issues you may encounter when transitioning from legacy to consolidated tools:

1. **Missing `operation` parameter**: All consolidated tools require an operation parameter
2. **Invalid operation names**: Check the correct operation name in the consolidated tool documentation
3. **Parameter format differences**: Some parameter formats differ between legacy and consolidated tools
4. **Validation rules**: Consolidated tools may have stricter validation rules

## Best Practices

1. **Always include the operation parameter** with consolidated tools
2. **Use the correct parameter structure** as specified in each consolidated tool's documentation
3. **Review error messages carefully** if you encounter issues during transition
4. **Check parameter validation requirements** for each operation

## Migration Example

### Legacy (Deprecated) Tools Approach:

```json
// Using separate tools for each operation
// Tool: calendar_list_events
{
  "startDateTime": "2025-05-15T00:00:00",
  "endDateTime": "2025-05-15T23:59:59",
  "maxResults": 10
}

// Tool: calendar_create_event
{
  "subject": "Team Meeting",
  "startDateTime": "2025-05-16T14:00:00",
  "endDateTime": "2025-05-16T15:00:00",
  "location": "Conference Room A",
  "attendees": ["john@example.com", "sarah@example.com"]
}
```

### Consolidated (Current) Tools Approach:

```json
// Same operations using the consolidated tool with operation parameter
// Tool: calendar
{
  "operation": "list",
  "startDateTime": "2025-05-15T00:00:00",
  "endDateTime": "2025-05-15T23:59:59",
  "maxResults": 10
}

// Tool: calendar
{
  "operation": "create",
  "subject": "Team Meeting",
  "startDateTime": "2025-05-16T14:00:00",
  "endDateTime": "2025-05-16T15:00:00",
  "location": "Conference Room A",
  "attendees": ["john@example.com", "sarah@example.com"]
}
```

## Conclusion

The tool consolidation provides a more manageable interface while maintaining the full functionality of the original tools. By using the operation parameter, you can perform all the same actions as before with fewer tools to navigate through.

If you encounter any issues with the consolidated tools, please refer to the ISSUE_RESOLUTION_SUMMARY.md and FIXES.md files for common problems and their solutions.