# Office MCP Server - Available Tools Documentation

This document provides comprehensive documentation for all available tools in the Office MCP Server, including usage examples for each tool.

## Complete Tool Count by Module

- **Authentication Tools**: 3 tools
- **Email Tools**: 5 consolidated tools
- **Calendar Tools**: 1 consolidated tool
- **Teams Tools**: 3 consolidated tools
- **Planner Tools**: 8 consolidated tools
- **Notification Tools**: 4 tools

**Total**: 24 tools available (consolidated from many original tools)

> **Note**: The server has been significantly consolidated to reduce the number of tools while maintaining full functionality. Individual tools have been replaced with operation-based tools.

## Table of Contents
- [Authentication Tools](#authentication-tools)
- [Email Tools](#email-tools)
- [Calendar Tools](#calendar-tools)
- [Teams Tools](#teams-tools)
- [Drive Tools](#drive-tools)
- [Planner Tools](#planner-tools)
- [User Management Tools](#user-management-tools)
- [Notification Tools](#notification-tools)

---

## Authentication Tools

### authenticate
Authenticate with Microsoft 365 using OAuth 2.0.

**Usage:**
```
Tool: authenticate
```

**Example:**
```
User: Please authenticate with Microsoft 365
Assistant: I'll authenticate with Microsoft 365 for you.

Tool: authenticate
Result: Please go to https://microsoft.com/devicelogin and enter code: ABC123
After authentication, you'll have access to all Microsoft 365 services.
```

### check_auth_status
Check the current authentication status.

**Usage:**
```
Tool: check_auth_status
```

**Example:**
```
User: Am I still logged in?
Assistant: Let me check your authentication status.

Tool: check_auth_status
Result: Authenticated as john.doe@company.com. Token expires in 59 minutes.
```

### logout
Log out from Microsoft 365.

**Usage:**
```
Tool: logout
```

---

## Email Tools

### email
Unified email tool for list, read, and send operations.

**Parameters:**
- `operation`: The operation to perform ("list", "read", or "send")
- Additional parameters based on the operation:
  
**list operation parameters:**
- `folderId` (optional): Folder ID to list emails from
- `maxResults` (optional): Maximum number of emails to retrieve (default: 10)

**read operation parameters:**
- `emailId`: Email message ID to read

**send operation parameters:**
- `to`: Array of recipient email addresses
- `subject`: Email subject
- `body`: Email body
- `cc` (optional): Array of CC recipients
- `bcc` (optional): Array of BCC recipients

**Example - List Emails:**
```
User: Show me my latest emails
Assistant: I'll list your recent emails.

Tool: email
Parameters: { 
  "operation": "list",
  "maxResults": 10 
}
Result: Found 10 emails:
- Meeting Tomorrow (from: boss@company.com)
- Project Update (from: colleague@company.com)
...
```

**Example - Read Email:**
```
User: Read the email with ID abc123
Assistant: I'll read that email for you.

Tool: email
Parameters: { 
  "operation": "read",
  "emailId": "abc123" 
}
Result: Subject: Project Status Update
From: team.lead@company.com
Content: The project is on track...
```

**Example - Send Email:**
```
User: Send an email to john@company.com about the meeting
Assistant: I'll send that email for you.

Tool: email
Parameters: {
  "operation": "send",
  "to": ["john@company.com"],
  "subject": "Meeting Update",
  "body": "Hi John, Just wanted to confirm our meeting tomorrow at 2 PM."
}
Result: Email sent successfully!
```

**Email Payload Format:**
```javascript
{
  message: {
    subject: "Email subject",
    body: {
      contentType: "HTML",  // or "Text"
      content: "Email body content"
    },
    toRecipients: [
      {
        emailAddress: {
          address: "recipient@example.com"
        }
      }
    ]
  },
  saveToSentItems: true  // Required parameter
}
```

This format follows the Microsoft Graph API requirements for the sendMail endpoint.

### email_search
Search emails with different modes: basic, enhanced, or simple.

**Parameters:**
- `query`: Search query string
- `mode` (optional): Search mode ("basic", "enhanced", "simple" - default: "basic")
- `maxResults` (optional): Maximum number of results (default: 10)
- Mode-specific parameters:
  - `from` (basic mode): Filter by sender email address
  - `subject` (basic mode): Filter by subject
  - `filterType` (simple mode): Type of filter ("subject", "from", "body" - default: "subject")

**Example:**
```
User: Find emails about the budget report
Assistant: I'll search for emails about the budget report.

Tool: email_search
Parameters: { "query": "budget report", "maxResults": 5 }
Result: Found 3 emails matching "budget report":
- Q4 Budget Report (from: finance@company.com)
...
```

### email_move
Move emails to a folder with optional batch processing.

**Parameters:**
- `emailIds`: Array of email IDs to move
- `destinationFolderId`: Destination folder ID
- `batch` (optional): Use batch processing (default: auto-enabled for >5 emails)

**Example:**
```
User: Move these 3 emails to the Archive folder
Assistant: I'll move those emails for you.

Tool: email_move
Parameters: {
  "emailIds": ["id1", "id2", "id3"],
  "destinationFolderId": "archiveFolderId"
}
Result: Moved 3 emails successfully.
```

### email_folder
Manage email folders: list or create.

**Parameters:**
- `operation`: The operation to perform ("list" or "create")
- Additional parameters based on operation:
  
**create operation parameters:**
- `displayName`: Folder display name
- `parentFolderId` (optional): Parent folder ID for creating subfolders

**Example:**
```
User: Show me my email folders
Assistant: I'll list your email folders.

Tool: email_folder
Parameters: { "operation": "list" }
Result: Found 12 email folders:
- Inbox (15 unread)
- Sent Items
...
```

### email_rules
Manage email rules: list or create.

**Parameters:**
- `operation`: The operation to perform ("list" or "create")
- `enhanced` (optional): Use enhanced mode for more features (default: false)
- Additional parameters for create operation:
  - `displayName`: Rule display name
  - `fromAddresses`: Array of sender email addresses to filter
  - `moveToFolder`: Folder ID to move matching emails to
  - `forwardTo`: Array of email addresses to forward to
  - Additional parameters for enhanced mode:
    - `subjectContains`: Array of subject keywords
    - `importance`: Importance level ("low", "normal", "high")

**Example:**
```
User: List my email rules
Assistant: I'll list your email rules.

Tool: email_rules
Parameters: { "operation": "list" }
Result: Found 2 email rules:
1. Move newsletters to Newsletters folder
   Enabled: true
...
```

---

## Calendar Tools

### calendar
Unified calendar tool for all calendar operations.

**Parameters:**
- `operation`: The operation to perform ("list", "create", "get", "update", or "delete")
- Additional parameters based on the operation:

**list operation parameters:**
- `startDateTime` (optional): ISO 8601 start date/time
- `endDateTime` (optional): ISO 8601 end date/time
- `maxResults` (optional): Maximum number of events (default: 10)

**create operation parameters:**
- `subject`: Event subject
- `start`: ISO 8601 start date/time
- `end`: ISO 8601 end date/time
- `location` (optional): Event location
- `content` (optional): Event description in HTML
- `attendees` (optional): Array of attendee emails
- `isOnlineMeeting` (optional): Create as Teams meeting (boolean)

**get operation parameters:**
- `eventId`: Event ID to retrieve

**update operation parameters:**
- `eventId`: Event ID to update
- `subject` (optional): New event subject
- `start` (optional): New ISO 8601 start date/time
- `end` (optional): New ISO 8601 end date/time
- `location` (optional): New event location
- `body` (optional): New event description

**delete operation parameters:**
- `eventId`: Event ID to delete
- `sendCancellations` (optional): Send cancellation notices (default: true)

**Example - List Calendar Events:**
```
User: What's on my calendar today?
Assistant: I'll check your calendar for today.

Tool: calendar
Parameters: {
  "operation": "list",
  "startDateTime": "2024-03-15T00:00:00Z",
  "endDateTime": "2024-03-15T23:59:59Z"
}
Result: Today's events:
- 9:00 AM: Team Standup
- 2:00 PM: Project Review
```

**Example - Create Calendar Event:**
```
User: Schedule a meeting with Sarah tomorrow at 3 PM
Assistant: I'll create that meeting for you.

Tool: calendar
Parameters: {
  "operation": "create",
  "subject": "Meeting with Sarah",
  "start": "2024-03-16T15:00:00Z",
  "end": "2024-03-16T16:00:00Z",
  "attendees": ["sarah@company.com"]
}
Result: Event created successfully! Event ID: xyz789
```

**Example - Update Calendar Event:**
```
User: Update my meeting with Sarah to 4 PM
Assistant: I'll update that meeting for you.

Tool: calendar
Parameters: {
  "operation": "update",
  "eventId": "xyz789",
  "start": "2024-03-16T16:00:00Z",
  "end": "2024-03-16T17:00:00Z"
}
Result: Event updated successfully!
```

**Example - Delete Calendar Event:**
```
User: Cancel my meeting with Sarah
Assistant: I'll cancel that meeting for you.

Tool: calendar
Parameters: {
  "operation": "delete",
  "eventId": "xyz789"
}
Result: Event deleted successfully!
```

---

## Teams Tools

The Teams module has been consolidated to three main operation-based tools, plus several legacy tools:

### teams_meeting
Unified tool for all Teams meeting operations.

**Parameters:**
- `operation`: The operation to perform
  - Available operations: "create", "update", "cancel", "get", "find_by_url", "list_transcripts", "get_transcript", "list_recordings", "get_recording", "get_participants", "get_insights"
- Additional parameters based on the operation:

**create operation parameters:**
- `subject`: Meeting subject
- `startDateTime`: ISO 8601 start date/time
- `endDateTime`: ISO 8601 end date/time
- `description` (optional): Meeting description
- `participants` (optional): Array of attendee emails

**Example - Create Meeting:**
```
User: Create a Teams meeting for tomorrow at 2 PM
Assistant: I'll create that Teams meeting.

Tool: teams_meeting
Parameters: {
  "operation": "create",
  "subject": "Project Discussion",
  "startDateTime": "2024-03-16T14:00:00Z",
  "endDateTime": "2024-03-16T15:00:00Z"
}
Result: Meeting created successfully!
Meeting ID: meet123
Join URL: https://teams.microsoft.com/l/meetup-join/...
```

**Example - Get Transcript:**
```
User: Get the transcript for the meeting
Assistant: I'll get that transcript for you.

Tool: teams_meeting
Parameters: {
  "operation": "list_transcripts",
  "meetingId": "meet123"
}
Result: Found 1 transcript:
- ID: transcript123
  Created: 03/16/2024, 3:00:00 PM
```

### teams_channel
Unified tool for all Teams channel operations.

**Parameters:**
- `operation`: The operation to perform
  - Available operations: "list", "create", "get", "update", "delete", "list_messages", "get_message", "create_message", "reply_to_message", "list_members", "add_member", "remove_member", "list_tabs"

**Example - List Channels:**
```
User: Show me the channels in the Engineering Team
Assistant: I'll list the channels in the Engineering Team.

Tool: teams_channel
Parameters: {
  "operation": "list",
  "teamId": "team123"
}
Result: Found 4 channels:
- General (ID: ch001)
- Development (ID: ch002)
...
```

**Example - Create Message:**
```
User: Send "Good morning team!" to the General channel
Assistant: I'll send that message to the General channel.

Tool: teams_channel
Parameters: {
  "operation": "create_message",
  "teamId": "team123",
  "channelId": "ch001",
  "content": "Good morning team!"
}
Result: Message created successfully! Message ID: msg789
```

### teams_chat
Unified tool for all Teams chat operations.

**Parameters:**
- `operation`: The operation to perform
  - Available operations: "list", "create", "get", "update", "delete", "list_messages", "get_message", "send_message", "update_message", "delete_message", "list_members", "add_member", "remove_member"

**Example - List Chats:**
```
User: Show me my Teams chats
Assistant: I'll list your Teams chats.

Tool: teams_chat
Parameters: {
  "operation": "list"
}
Result: Found 5 chats:
1. Marketing Team Group Chat
   Type: Group Chat
   Last updated: 3/15/2024, 2:30:00 PM
   ID: chat001
...
```

**Example - Send Chat Message:**
```
User: Send a message to John in Teams
Assistant: I'll send that message for you.

Tool: teams_chat
Parameters: {
  "operation": "send_message",
  "chatId": "chat123",
  "content": "Hi John, are we still meeting at 2 PM?"
}
Result: Message sent successfully! Message ID: msg456
```


---

## Drive Tools

### drive_list_files
List files in OneDrive or SharePoint.

**Parameters:**
- `folderId` (optional): Folder ID to list files from
- `driveId` (optional): Drive ID

**Example:**
```
User: Show me files in my OneDrive
Assistant: I'll list your OneDrive files.

Tool: drive_list_files
Result: Found 15 items:
- Presentation.pptx [File] (125 KB) - ID: file001
- Projects [Folder] - ID: folder001
...
```

### drive_upload_file
Upload a file to OneDrive.

**Parameters:**
- `filePath`: Local path to the file
- `destinationPath`: Destination path in OneDrive
- `driveId` (optional): Drive ID
- `folderId` (optional): Destination folder ID

**Example:**
```
User: Upload report.pdf to my OneDrive
Assistant: I'll upload report.pdf to your OneDrive.

Tool: drive_upload_file
Parameters: {
  "filePath": "/local/path/report.pdf",
  "destinationPath": "Documents"
}
Result: File uploaded successfully!
File ID: file789
```

### drive_create_folder
Create a new folder.

**Parameters:**
- `folderName`: Name of the folder
- `parentPath` (optional): Parent path
- `driveId` (optional): Drive ID
- `parentFolderId` (optional): Parent folder ID

**Example:**
```
User: Create a new folder called "2024 Reports"
Assistant: I'll create that folder for you.

Tool: drive_create_folder
Parameters: {
  "folderName": "2024 Reports"
}
Result: Folder created successfully!
Folder ID: folder789
```

### Additional Drive Tools

The Drive module also includes these additional tools:

- **drive_get_file_content**: Get file content
- **drive_search_files**: Search for files
- **drive_copy_item**: Copy a file or folder
- **drive_move_item**: Move a file or folder
- **drive_delete_item**: Delete a file or folder
- **drive_list_recycle_bin**: List items in the recycle bin
- **drive_restore_item**: Restore an item from the recycle bin
- **drive_permanently_delete**: Permanently delete an item from the recycle bin
- **drive_empty_recycle_bin**: Empty the entire recycle bin

---

## Planner Tools

### planner_list_plans
List all Planner plans for the current user.

**Example:**
```
User: Show me my Planner plans
Assistant: I'll list your Planner plans.

Tool: planner_list_plans
Result: Found 3 plans:
- Q1 Marketing Plan (ID: plan001)
- Product Launch (ID: plan002)
...
```

### planner_create_task
Create a new task in a Planner plan.

**Parameters:**
- `planId`: Plan ID
- `title`: Task title
- `bucketId` (optional): Bucket ID
- `dueDateTime` (optional): Due date in ISO 8601 format
- `assignedTo` (optional): User ID or array of user IDs
- `percentComplete` (optional): Completion percentage (0-100)

**Example:**
```
User: Create a task "Review documentation" in the Product Launch plan
Assistant: I'll create that task in your plan.

Tool: planner_create_task
Parameters: {
  "planId": "plan002",
  "title": "Review documentation",
  "dueDateTime": "2024-03-20T17:00:00Z"
}
Result: Task created successfully! Task ID: task123
```

### planner_update_task
Update an existing task.

**Parameters:**
- `taskId`: Task ID
- `title` (optional): New task title
- `percentComplete` (optional): Completion percentage
- `dueDateTime` (optional): Due date in ISO 8601 format
- `bucketId` (optional): Bucket ID

**Example:**
```
User: Mark task123 as 50% complete
Assistant: I'll update that task's completion status.

Tool: planner_update_task
Parameters: {
  "taskId": "task123",
  "percentComplete": 50
}
Result: Task updated successfully!
```

### Additional Planner Tools

The Planner module also includes these additional tools:

- **planner_get_plan**: Get details of a specific plan
- **planner_create_plan**: Create a new plan
- **planner_update_plan**: Update plan details
- **planner_delete_plan**: Delete a plan
- **planner_list_tasks**: List tasks in a plan
- **planner_get_task**: Get task details
- **planner_delete_task**: Delete a task
- **planner_assign_task**: Assign a task to a user
- **planner_list_buckets**: List buckets in a plan
- **planner_create_bucket**: Create a new bucket
- **planner_update_bucket**: Update bucket details
- **planner_delete_bucket**: Delete a bucket
- **planner_get_bucket_tasks**: Get tasks in a specific bucket

---

## User Management Tools

### user_get_current_profile
Get the current user's profile information.

**Example:**
```
User: Show me my profile information
Assistant: I'll retrieve your profile information.

Tool: user_get_current_profile
Result: User Profile:
Name: John Doe
Email: john.doe@company.com
Job Title: Software Engineer
Department: Engineering
```

### user_search
Search for users in the organization directory.

**Parameters:**
- `searchQuery` (optional): Search query
- `filter` (optional): OData filter expression
- `top` (optional): Maximum number of results

**Example:**
```
User: Find users named Sarah
Assistant: I'll search for users named Sarah.

Tool: user_search
Parameters: { "searchQuery": "Sarah" }
Result: Found 3 users:
- Sarah Johnson (sarah.johnson@company.com) - Marketing Manager
- Sarah Smith (sarah.smith@company.com) - Developer
...
```

### user_get_presence
Get a user's presence information.

**Parameters:**
- `userId` (optional): User ID or email (defaults to current user)

**Example:**
```
User: Is Sarah online?
Assistant: I'll check Sarah's presence status.

Tool: user_get_presence
Parameters: { "userId": "sarah.johnson@company.com" }
Result: User Presence:
Availability: Available
Activity: Available
Status Message: None
In Meeting: No
```

### user_set_presence
Set the current user's presence.

**Parameters:**
- `availability`: Availability status (Available, Busy, DoNotDisturb, etc.)
- `activity`: Activity type (Available, InACall, InAMeeting, etc.)
- `expirationMinutes` (optional): Minutes until the presence expires

**Example:**
```
User: Set my status to busy for the next hour
Assistant: I'll set your status to busy for the next hour.

Tool: user_set_presence
Parameters: {
  "availability": "Busy",
  "activity": "Busy",
  "expirationMinutes": 60
}
Result: Presence updated successfully!
```

### Additional User Management Tools

The User Management module also includes these additional tools:

- **user_get_profile**: Get a specific user's profile
- **user_update_current_profile**: Update current user's profile
- **user_get_manager**: Get a user's manager information
- **user_get_direct_reports**: Get a user's direct reports
- **user_get_photo**: Get a user's profile photo
- **user_list**: List all users in the organization
- **user_get_by_department**: Get users by department
- **user_get_by_location**: Get users by office location
- **user_list_groups**: List organization groups
- **user_set_status_message**: Set status message
- **user_clear_status_message**: Clear status message
- **user_get_bulk_presence**: Get presence for multiple users

---

## Notification Tools

### notification_create_subscription
Create a subscription for change notifications.

**Parameters:**
- `resource`: Resource to monitor
- `changeType`: Type of change to monitor
- `notificationUrl`: URL to receive notifications
- `expirationMinutes` (optional): Expiration time in minutes
- `includeResourceData` (optional): Include resource data in notifications

**Example:**
```
User: Subscribe to email notifications
Assistant: I'll create a subscription for email notifications.

Tool: notification_create_subscription
Parameters: {
  "resource": "/me/messages",
  "changeType": "created",
  "notificationUrl": "https://webhook.site/abc123",
  "expirationMinutes": 60
}
Result: Subscription created successfully!
ID: sub123
```

### Additional Notification Tools

The Notifications module also includes these additional tools:

- **notification_list_subscriptions**: List existing subscriptions
- **notification_get_subscription**: Get details of a specific subscription
- **notification_update_subscription**: Update a subscription
- **notification_delete_subscription**: Delete a subscription

---

## Error Handling

All tools handle common errors gracefully:

- **Authentication errors**: Will prompt to authenticate first
- **Missing parameters**: Will indicate which parameters are required
- **API errors**: Will provide descriptive error messages

## Best Practices

1. **Authentication**: Always ensure authentication before using any tools
2. **Parameters**: Provide all required parameters for successful operations
3. **Date/Time**: Use ISO 8601 format for all date/time parameters
4. **IDs**: Store returned IDs for subsequent operations
5. **Limits**: Be mindful of API limits when retrieving large datasets

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Office MCP Server README](./README.md)
- [Authentication Guide](./auth/README.md)
