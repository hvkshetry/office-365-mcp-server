/**
 * Consolidated Calendar module
 * Reduces from 5 tools to 1 tool with operation parameters
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Unified calendar handler for all calendar operations
 */
async function handleCalendar(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create, get, update, delete" 
      }]
    };
  }
  
  try {
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listCalendarEvents(accessToken, params);
      case 'create':
        return await createCalendarEvent(accessToken, params);
      case 'get':
        return await getCalendarEvent(accessToken, params);
      case 'update':
        return await updateCalendarEvent(accessToken, params);
      case 'delete':
        return await deleteCalendarEvent(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create, get, update, delete` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in calendar ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in calendar ${operation}: ${error.message}` }]
    };
  }
}

// Implementation functions (existing logic from original files)

async function listCalendarEvents(accessToken, params) {
  const { startDateTime, endDateTime, maxResults = 10 } = params;
  
  console.error(`Calendar list request with startDateTime: ${startDateTime}, endDateTime: ${endDateTime}`);
  
  const queryParams = {
    $select: config.CALENDAR_SELECT_FIELDS,
    $orderby: 'start/dateTime',
    $top: maxResults
  };
  
  // Add date filter if provided
  if (startDateTime || endDateTime) {
    const filters = [];
    if (startDateTime) {
      // Format the date properly for Microsoft Graph API
      // The API expects dates in ISO 8601 format
      const formattedStartDateTime = startDateTime.includes('Z') 
        ? startDateTime 
        : `${startDateTime}Z`;
      
      filters.push(`start/dateTime ge '${formattedStartDateTime}'`);
      console.error(`Using start filter: start/dateTime ge '${formattedStartDateTime}'`);
    }
    if (endDateTime) {
      // Format the date properly for Microsoft Graph API
      const formattedEndDateTime = endDateTime.includes('Z') 
        ? endDateTime 
        : `${endDateTime}Z`;
      
      filters.push(`end/dateTime le '${formattedEndDateTime}'`);
      console.error(`Using end filter: end/dateTime le '${formattedEndDateTime}'`);
    }
    queryParams.$filter = filters.join(' and ');
  }
  
  // If requesting events by date, add a fallback filter on view range
  if (startDateTime && endDateTime) {
    // Try calendar view approach as a more reliable way to get events in a specific date range
    try {
      console.error('Attempting to use calendarView endpoint for more reliable results');
      const formattedStartDateTime = startDateTime.includes('Z') ? startDateTime : `${startDateTime}Z`;
      const formattedEndDateTime = endDateTime.includes('Z') ? endDateTime : `${endDateTime}Z`;
      
      const viewParams = {
        $select: config.CALENDAR_SELECT_FIELDS,
        $orderby: 'start/dateTime',
        $top: maxResults
      };
      
      const viewResponse = await callGraphAPI(
        accessToken,
        'GET',
        'me/calendar/calendarView',
        null,
        {
          ...viewParams,
          startDateTime: formattedStartDateTime,
          endDateTime: formattedEndDateTime
        }
      );
      
      if (viewResponse.value && viewResponse.value.length > 0) {
        console.error(`calendarView returned ${viewResponse.value.length} events`);
        return formatCalendarResponse(viewResponse);
      }
      console.error('calendarView returned no events, falling back to regular query');
    } catch (viewError) {
      console.error('Error using calendarView approach, falling back to regular query:', viewError);
    }
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/calendar/events',
    null,
    queryParams
  );
  
  return formatCalendarResponse(response);
}

/**
 * Helper function to format calendar response
 */
function formatCalendarResponse(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No calendar events found." }]
    };
  }
  
  const eventsList = response.value.map((event, index) => {
    const startTime = new Date(event.start.dateTime);
    const endTime = new Date(event.end.dateTime);
    return `${index + 1}. ${event.subject}\n   Start: ${startTime.toLocaleString()}\n   End: ${endTime.toLocaleString()}\n   Location: ${event.location.displayName || 'N/A'}\n   ID: ${event.id}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} calendar events:\n\n${eventsList}` 
    }]
  }
}

async function createCalendarEvent(accessToken, params) {
  const { subject, content, start, end, location, attendees, isOnlineMeeting } = params;
  
  if (!subject || !start || !end) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters: subject, start, and end" 
      }]
    };
  }
  
  const event = {
    subject: subject,
    body: {
      contentType: "HTML",
      content: content || ""
    },
    start: {
      dateTime: start,
      timeZone: "UTC"
    },
    end: {
      dateTime: end,
      timeZone: "UTC"
    }
  };
  
  if (location) {
    event.location = { displayName: location };
  }
  
  if (attendees && attendees.length > 0) {
    event.attendees = attendees.map(email => ({
      emailAddress: { address: email },
      type: "required"
    }));
  }
  
  if (isOnlineMeeting) {
    event.isOnlineMeeting = true;
    event.onlineMeetingProvider = "teamsForBusiness";
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/calendar/events',
    event
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Calendar event created successfully!\nEvent ID: ${response.id}` 
    }]
  };
}

async function getCalendarEvent(accessToken, params) {
  const { eventId } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `me/calendar/events/${eventId}`,
    null,
    {
      $select: config.CALENDAR_SELECT_FIELDS
    }
  );
  
  const startTime = new Date(response.start.dateTime);
  const endTime = new Date(response.end.dateTime);
  
  let eventDetails = `Subject: ${response.subject}\n`;
  eventDetails += `Start: ${startTime.toLocaleString()}\n`;
  eventDetails += `End: ${endTime.toLocaleString()}\n`;
  eventDetails += `Location: ${response.location?.displayName || 'N/A'}\n`;
  eventDetails += `Online Meeting: ${response.isOnlineMeeting ? 'Yes' : 'No'}\n`;
  
  if (response.attendees && response.attendees.length > 0) {
    eventDetails += `Attendees: ${response.attendees.map(a => a.emailAddress.address).join(', ')}\n`;
  }
  
  if (response.body?.content) {
    eventDetails += `\nDescription:\n${response.body.content}`;
  }
  
  return {
    content: [{ type: "text", text: eventDetails }]
  };
}

async function updateCalendarEvent(accessToken, params) {
  const { eventId, ...updateFields } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  const allowedFields = ['subject', 'location', 'body', 'start', 'end', 'isOnlineMeeting'];
  const update = {};
  
  // Build update object with allowed fields
  for (const [key, value] of Object.entries(updateFields)) {
    if (allowedFields.includes(key)) {
      if (key === 'location') {
        update.location = { displayName: value };
      } else if (key === 'body') {
        update.body = { contentType: "HTML", content: value };
      } else if (key === 'start' || key === 'end') {
        update[key] = { dateTime: value, timeZone: "UTC" };
      } else {
        update[key] = value;
      }
    }
  }
  
  if (Object.keys(update).length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "No valid fields to update. Valid fields: subject, location, body, start, end, isOnlineMeeting" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `me/calendar/events/${eventId}`,
    update
  );
  
  return {
    content: [{ type: "text", text: "Calendar event updated successfully!" }]
  };
}

async function deleteCalendarEvent(accessToken, params) {
  const { eventId, sendCancellations = true } = params;
  
  if (!eventId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: eventId" 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `me/calendar/events/${eventId}`,
    null,
    sendCancellations ? { sendCancellations: 'true' } : {}
  );
  
  return {
    content: [{ type: "text", text: "Calendar event deleted successfully!" }]
  };
}

// Export consolidated tool
const calendarTools = [
  {
    name: "calendar",
    description: "Manage calendar events: list, create, get, update, or delete",
    inputSchema: {
      type: "object",
      properties: {
        operation: { 
          type: "string", 
          enum: ["list", "create", "get", "update", "delete"],
          description: "The operation to perform" 
        },
        // List parameters
        startDateTime: { type: "string", description: "Start date/time filter in ISO format (for list)" },
        endDateTime: { type: "string", description: "End date/time filter in ISO format (for list)" },
        maxResults: { type: "number", description: "Maximum number of results (default: 10)" },
        // Create parameters
        subject: { type: "string", description: "Event subject (for create/update)" },
        content: { type: "string", description: "Event body content in HTML (for create/update)" },
        start: { type: "string", description: "Start date/time in ISO format (for create/update)" },
        end: { type: "string", description: "End date/time in ISO format (for create/update)" },
        location: { type: "string", description: "Event location (for create/update)" },
        attendees: { 
          type: "array", 
          items: { type: "string" },
          description: "Attendee email addresses (for create)" 
        },
        isOnlineMeeting: { type: "boolean", description: "Create as Teams meeting (for create/update)" },
        // Get/Update/Delete parameters
        eventId: { type: "string", description: "Event ID (for get/update/delete)" },
        // Delete parameters
        sendCancellations: { type: "boolean", description: "Send cancellation notices (default: true)" }
      },
      required: ["operation"]
    },
    handler: handleCalendar
  }
];

module.exports = { calendarTools };