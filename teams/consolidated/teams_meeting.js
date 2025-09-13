/**
 * Consolidated Teams Meeting Module
 * 
 * Provides a unified interface for all Teams meeting operations including:
 * - Creating/updating/canceling meetings
 * - Finding meeting details
 * - Accessing meeting transcripts
 * - Getting meeting insights
 * - Working with meeting recordings
 * 
 * Teams Meeting ID Format Notes:
 * 
 * Microsoft Teams uses different ID formats:
 * 
 * 1. Online Meeting ID - Used for transcript and recording APIs
 *    Format: MSo1N2Y5ZGFjYy03MWJmLTQ3NDMtYjQxMy01M2EdFGkdRWHJlQ
 * 
 * 2. Thread ID - Used for chat and channel messages
 *    Format: 19:meeting_ZWQxMjQ1OTUtNzY4ZC00Y2FmLTg4ZTQtYjRkNDIyY2NmZTZi@thread.v2
 * 
 * 3. Join URL - Contains meeting ID in the path
 *    Format: https://teams.microsoft.com/l/meetup-join/...
 * 
 * The transcript API requires the Online Meeting ID format.
 * If a Thread ID is provided, we attempt to convert it to a Meeting ID.
 * 
 * Troubleshooting:
 * - Make sure the meeting has ended (transcripts are only available after meeting end)
 * - Make sure transcription was enabled during the meeting
 * - Only meeting organizers or those with proper permissions can access transcripts
 * - The meeting ID must be in the correct format
 */

const { ensureAuthenticated } = require('../../auth');
const { callGraphAPI } = require('../../utils/graph-api');
const config = require('../../config');

/**
 * Main handler for teams_meeting operations
 * @param {Object} args - The operation arguments
 * @returns {Object} - MCP response
 */
async function handleTeamsMeeting(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: create, update, cancel, get, find_by_url, list_transcripts, get_transcript, list_recordings, get_recording, get_participants, get_insights" 
      }]
    };
  }
  
  try {
    console.error(`teams_meeting operation: ${operation}`);
    console.error('teams_meeting params:', JSON.stringify(params));
    
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'create':
        return await createMeeting(accessToken, params);
      case 'update':
        return await updateMeeting(accessToken, params);
      case 'cancel':
        return await cancelMeeting(accessToken, params);
      case 'get':
        return await getMeeting(accessToken, params);
      case 'find_by_url':
        return await findMeetingByUrl(accessToken, params);
      case 'list_transcripts':
        return await listTranscripts(accessToken, params);
      case 'get_transcript':
        return await getTranscript(accessToken, params);
      case 'list_recordings':
        return await listRecordings(accessToken, params);
      case 'get_recording':
        return await getRecording(accessToken, params);
      case 'get_participants':
        return await getMeetingParticipants(accessToken, params);
      case 'get_insights':
        return await getMeetingInsights(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: create, update, cancel, get, find_by_url, list_transcripts, get_transcript, list_recordings, get_recording, get_participants, get_insights` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in teams_meeting ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in teams_meeting operation: ${error.message}` }]
    };
  }
}

/**
 * Create a Teams online meeting
 */
async function createMeeting(accessToken, params) {
  const { subject, startDateTime, endDateTime, description, participants } = params;
  
  if (!subject || !startDateTime || !endDateTime) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide subject, startDateTime, and endDateTime." 
      }]
    };
  }
  
  const meeting = {
    subject,
    startDateTime,
    endDateTime,
    participants: {}
  };
  
  if (description) {
    meeting.description = description;
  }
  
  if (participants) {
    // Format participants if provided
    meeting.participants = {
      attendees: []
    };
    
    if (Array.isArray(participants)) {
      meeting.participants.attendees = participants.map(email => ({
        upn: email,
        role: 'attendee'
      }));
    }
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    'me/onlineMeetings',
    meeting
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Meeting created successfully!\nMeeting ID: ${response.id}\nJoin URL: ${response.joinWebUrl}` 
    }]
  };
}

/**
 * Update an existing meeting
 */
async function updateMeeting(accessToken, params) {
  const { meetingId, subject, startDateTime, endDateTime, description } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  const updateData = {};
  if (subject) updateData.subject = subject;
  if (startDateTime) updateData.startDateTime = startDateTime;
  if (endDateTime) updateData.endDateTime = endDateTime;
  if (description) updateData.description = description;
  
  if (Object.keys(updateData).length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: "No update parameters provided. Please specify at least one field to update." 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'PATCH',
    `me/onlineMeetings/${meetingId}`,
    updateData
  );
  
  return {
    content: [{ type: "text", text: "Meeting updated successfully!" }]
  };
}

/**
 * Cancel an existing meeting
 */
async function cancelMeeting(accessToken, params) {
  const { meetingId, comment } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  // For online meetings, deleting is equivalent to cancelling
  await callGraphAPI(
    accessToken,
    'DELETE',
    `me/onlineMeetings/${meetingId}`
  );
  
  return {
    content: [{ type: "text", text: "Meeting cancelled successfully!" }]
  };
}

/**
 * Get meeting details by ID
 */
async function getMeeting(accessToken, params) {
  const { meetingId } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  const meeting = await callGraphAPI(
    accessToken,
    'GET',
    `me/onlineMeetings/${meetingId}`
  );
  
  let meetingInfo = `Meeting: ${meeting.subject}\n`;
  meetingInfo += `ID: ${meeting.id}\n`;
  meetingInfo += `Start: ${new Date(meeting.startDateTime).toLocaleString()}\n`;
  meetingInfo += `End: ${new Date(meeting.endDateTime).toLocaleString()}\n`;
  meetingInfo += `Join URL: ${meeting.joinWebUrl}\n`;
  meetingInfo += `Meeting Type: ${meeting.onlineMeetingType || 'Teams'}\n`;
  meetingInfo += `Audio Conference: ${meeting.audioConferencing ? 'Enabled' : 'Disabled'}\n`;
  
  if (meeting.participants) {
    const attendeeCount = meeting.participants.attendees?.length || 0;
    meetingInfo += `Participants: ${attendeeCount} attendees\n`;
  }
  
  return {
    content: [{ type: "text", text: meetingInfo }]
  };
}

/**
 * Find a meeting by its join URL
 */
async function findMeetingByUrl(accessToken, params) {
  const { joinUrl } = params;
  
  if (!joinUrl) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: joinUrl" 
      }]
    };
  }
  
  try {
    // URL encode the join URL since it's used in a filter
    const encodedUrl = encodeURIComponent(joinUrl);
    
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/onlineMeetings',
      null,
      {
        $filter: `joinWebUrl eq '${encodedUrl}'`
      }
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No meeting found with the provided join URL." }]
      };
    }
    
    const meeting = response.value[0];
    
    return await getMeeting(accessToken, { meetingId: meeting.id });
  } catch (error) {
    console.error('Error finding meeting by URL:', error);
    
    // Try to extract meeting ID from URL
    try {
      const meetingId = extractMeetingIdFromUrl(joinUrl);
      if (meetingId) {
        return {
          content: [{ 
            type: "text", 
            text: `Unable to find meeting directly by URL, but extracted potential meeting ID: ${meetingId}\n\nTry using this ID with the get operation.` 
          }]
        };
      }
    } catch (extractError) {
      console.error('Error extracting meeting ID from URL:', extractError);
    }
    
    throw error;
  }
}

/**
 * Extract meeting ID from join URL
 */
function extractMeetingIdFromUrl(meetingUrl) {
  try {
    const url = new URL(meetingUrl);
    const pathSegments = url.pathname.split('/');
    
    // Look for the meeting ID in the path (usually after 'meetup-join')
    const meetupIndex = pathSegments.findIndex(segment => segment === 'meetup-join');
    if (meetupIndex !== -1 && pathSegments[meetupIndex + 1]) {
      const meetingId = decodeURIComponent(pathSegments[meetupIndex + 1]);
      return meetingId;
    }
    
    // Check for meeting ID in query parameters
    const meetingIdParam = url.searchParams.get('meetingId');
    if (meetingIdParam) {
      return meetingIdParam;
    }
    
    return null;
  } catch (error) {
    console.error('Error extracting meeting ID from URL:', error);
    return null;
  }
}

/**
 * Convert a Teams thread ID to an online meeting ID using official Microsoft Graph approach
 * @param {string} accessToken - The access token for API calls
 * @param {string} threadId - The Teams thread ID (format: 19:meeting_XXX@thread.v2)
 * @returns {Promise<string>} - The online meeting ID
 */
async function convertThreadIdToMeetingId(accessToken, threadId) {
  try {
    console.error(`Converting thread ID to meeting ID: ${threadId}`);
    
    // Step 1: Get chat details to retrieve the join URL
    const chatDetails = await callGraphAPI(
      accessToken,
      'GET',
      `chats/${threadId}`
    );
    
    console.error('Chat details retrieved:', JSON.stringify(chatDetails, null, 2));
    
    // Check if this is a meeting chat with online meeting info
    if (!chatDetails.onlineMeetingInfo || !chatDetails.onlineMeetingInfo.joinWebUrl) {
      throw new Error(`Chat ${threadId} does not have online meeting information. This may not be a meeting chat.`);
    }
    
    const joinWebUrl = chatDetails.onlineMeetingInfo.joinWebUrl;
    console.error(`Found join URL: ${joinWebUrl}`);
    
    // Step 2: Use the join URL to find the online meeting
    const meetings = await callGraphAPI(
      accessToken,
      'GET',
      'me/onlineMeetings',
      null,
      {
        $filter: `joinWebUrl eq '${joinWebUrl}'`
      }
    );
    
    console.error('Meeting search results:', JSON.stringify(meetings, null, 2));
    
    if (!meetings.value || meetings.value.length === 0) {
      throw new Error(`No online meeting found for join URL: ${joinWebUrl}`);
    }
    
    const meetingId = meetings.value[0].id;
    console.error(`Successfully converted thread ID to meeting ID: ${meetingId}`);
    
    return meetingId;
    
  } catch (error) {
    console.error(`Error converting thread ID to meeting ID: ${error.message}`);
    
    // Provide specific error messages based on the failure type
    if (error.message.includes('does not have online meeting information')) {
      throw new Error(`The provided chat ID ${threadId} is not associated with a Teams meeting. Please ensure you're using a thread ID from a meeting chat.`);
    }
    
    if (error.message.includes('No online meeting found')) {
      throw new Error(`Unable to locate the online meeting. The meeting may have expired or you may not have access to it.`);
    }
    
    throw new Error(`Failed to convert thread ID to meeting ID: ${error.message}`);
  }
}

/**
 * List transcripts for a meeting
 */
async function listTranscripts(accessToken, params) {
  const { meetingId } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  try {
    // Determine the type of ID provided
    const isThreadId = meetingId.includes('@thread.v2');
    console.error(`Meeting ID format: ${isThreadId ? 'Thread ID' : 'Meeting ID'} - ${meetingId}`);
    
    // If it's a thread ID, we need to convert it to a meeting ID
    let actualMeetingId = meetingId;
    
    if (isThreadId) {
      try {
        console.error('Attempting to convert thread ID to meeting ID');
        console.error('Thread ID being converted:', meetingId);
        actualMeetingId = await convertThreadIdToMeetingId(accessToken, meetingId);
        console.error('Conversion successful. New meeting ID:', actualMeetingId);
      } catch (conversionError) {
        console.error(`Unable to convert thread ID to meeting ID: ${conversionError.message}`);
        console.error('Conversion error stack:', conversionError.stack);
        return {
          content: [{ 
            type: "text", 
            text: `DEBUG: Thread ID conversion failed. Original ID: ${meetingId}. Error: ${conversionError.message}` 
          }]
        };
      }
    }
    
    // Now use the proper meeting ID to get transcripts
    console.error(`Using meeting ID for transcript retrieval: ${actualMeetingId}`);
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${actualMeetingId}/transcripts`
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No transcripts found for this meeting. This could be because the meeting has not ended, transcription was not enabled, or no transcripts were generated." 
        }]
      };
    }
    
    // Format the transcript list
    const transcriptList = response.value.map(transcript => 
      `- ID: ${transcript.id}\n  Created: ${new Date(transcript.createdDateTime).toLocaleString()}`
    ).join('\n\n');
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} transcripts:\n\n${transcriptList}\n\nUse teams_meeting with operation=get_transcript to retrieve the content of a specific transcript.` 
      }]
    };
  } catch (error) {
    console.error(`Error in listTranscripts:`, error);
    
    // Provide helpful error messages based on the error type
    if (error.message && error.message.includes('Invalid meeting id')) {
      return {
        content: [{ 
          type: "text", 
          text: `Invalid meeting ID format. Please provide a valid online meeting ID, not a chat or thread ID. If you have a chat ID, you can use the meeting's join URL instead.` 
        }]
      };
    }
    
    if (error.message && error.message.includes('Access denied')) {
      return {
        content: [{ 
          type: "text", 
          text: `Access denied to meeting transcripts. You need to be the meeting organizer or have proper permissions to access transcripts.` 
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error retrieving transcripts: ${error.message}. Make sure the meeting ID is valid and that you have permission to access this meeting.` 
      }]
    };
  }
}

/**
 * Get transcript content
 */
async function getTranscript(accessToken, params) {
  const { meetingId, transcriptId, format = 'text/vtt' } = params;
  
  if (!meetingId || !transcriptId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide meetingId and transcriptId." 
      }]
    };
  }
  
  try {
    // If it's a thread ID, we need to convert it to a meeting ID
    let actualMeetingId = meetingId;
    
    if (meetingId.includes('@thread.v2')) {
      try {
        console.error('Attempting to convert thread ID to meeting ID for transcript content');
        actualMeetingId = await convertThreadIdToMeetingId(accessToken, meetingId);
      } catch (conversionError) {
        console.error(`Unable to convert thread ID to meeting ID: ${conversionError.message}`);
        return {
          content: [{ 
            type: "text", 
            text: `Unable to convert the Teams chat ID to a meeting ID. Please use the meeting ID directly or the meeting's join URL. Error: ${conversionError.message}` 
          }]
        };
      }
    }
    
    // First verify the transcript exists
    try {
      await callGraphAPI(
        accessToken,
        'GET',
        `me/onlineMeetings/${actualMeetingId}/transcripts/${transcriptId}`
      );
    } catch (error) {
      return {
        content: [{ 
          type: "text", 
          text: `Transcript not found or not accessible. Error: ${error.message}` 
        }]
      };
    }
    
    // Get the actual transcript content
    const transcriptContent = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${actualMeetingId}/transcripts/${transcriptId}/content`,
      null,
      {
        $format: format
      }
    );
    
    // Process the VTT content to a more readable format
    let processedContent = "Transcript content:\n\n";
    
    if (typeof transcriptContent === 'string') {
      // Basic formatting for VTT content
      try {
        const lines = transcriptContent.split('\n');
        let currentSpeaker = null;
        let currentUtterance = [];
        
        for (const line of lines) {
          // Skip WebVTT header and empty lines
          if (line === 'WEBVTT' || line.trim() === '') continue;
          
          // Check if line contains speaker information
          if (line.includes('-->')) {
            // This is a timestamp line, next line should be content
            continue;
          }
          
          // If we have a speaker and an utterance, add it to the result
          if (line.startsWith('<v ')) {
            // New speaker line
            if (currentSpeaker && currentUtterance.length > 0) {
              processedContent += `${currentSpeaker}: ${currentUtterance.join(' ')}\n\n`;
              currentUtterance = [];
            }
            
            // Extract speaker name
            const speakerMatch = line.match(/<v ([^>]+)>(.*)/);
            if (speakerMatch) {
              currentSpeaker = speakerMatch[1];
              currentUtterance.push(speakerMatch[2].trim());
            }
          } else if (currentSpeaker) {
            // Continue current utterance
            currentUtterance.push(line.trim());
          }
        }
        
        // Add the last utterance
        if (currentSpeaker && currentUtterance.length > 0) {
          processedContent += `${currentSpeaker}: ${currentUtterance.join(' ')}\n\n`;
        }
      } catch (parseError) {
        console.error('Error parsing transcript:', parseError);
        processedContent += transcriptContent;
      }
    } else {
      processedContent += 'Transcript content is not in expected format.';
    }
    
    return {
      content: [{ type: "text", text: processedContent }]
    };
  } catch (error) {
    console.error(`Error in getTranscript:`, error);
    
    // Provide helpful error messages based on the error type
    if (error.message && error.message.includes('Invalid meeting id')) {
      return {
        content: [{ 
          type: "text", 
          text: `Invalid meeting ID format. Please provide a valid online meeting ID, not a chat or thread ID. If you have a chat ID, you can use the meeting's join URL instead.` 
        }]
      };
    }
    
    if (error.message && error.message.includes('Access denied')) {
      return {
        content: [{ 
          type: "text", 
          text: `Access denied to meeting transcript. You need to be the meeting organizer or have proper permissions to access this transcript.` 
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `Error retrieving transcript: ${error.message}. Make sure the meeting ID and transcript ID are valid and that you have permission to access this meeting.` 
      }]
    };
  }
}

/**
 * List recordings for a meeting
 */
async function listRecordings(accessToken, params) {
  const { meetingId } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${meetingId}/recordings`
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No recordings found for this meeting." }]
      };
    }
    
    // Format the recording list
    const recordingList = response.value.map((recording, index) => {
      return `${index + 1}. Created: ${new Date(recording.createdDateTime).toLocaleString()}\n   ID: ${recording.id}\n   Duration: ${formatDuration(recording.meetingChatId || 'Unknown')}\n`;
    }).join('\n');
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} recordings:\n\n${recordingList}` 
      }]
    };
  } catch (error) {
    // Handle case where API might not be available yet
    return {
      content: [{ 
        type: "text", 
        text: `Unable to retrieve recordings. This might be because the recording is still processing or the API functionality is limited: ${error.message}` 
      }]
    };
  }
}

/**
 * Format duration in seconds to readable format
 */
function formatDuration(seconds) {
  if (isNaN(seconds)) return 'Unknown';
  
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remainingSeconds = Math.floor(seconds % 60);
  
  let result = '';
  if (hours > 0) result += `${hours}h `;
  if (minutes > 0 || hours > 0) result += `${minutes}m `;
  result += `${remainingSeconds}s`;
  
  return result.trim();
}

/**
 * Get recording content
 */
async function getRecording(accessToken, params) {
  const { meetingId, recordingId } = params;
  
  if (!meetingId || !recordingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide meetingId and recordingId." 
      }]
    };
  }
  
  try {
    const recording = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${meetingId}/recordings/${recordingId}`
    );
    
    let recordingInfo = `Recording ID: ${recording.id}\n`;
    recordingInfo += `Created: ${new Date(recording.createdDateTime).toLocaleString()}\n`;
    
    if (recording.recordingContentUrl) {
      recordingInfo += `\nDownload URL: ${recording.recordingContentUrl}\n`;
      recordingInfo += `\nNote: This URL is temporary and will expire. Download the recording promptly.`;
    } else {
      recordingInfo += `\nThis recording doesn't have a direct download URL available through the API.`;
    }
    
    return {
      content: [{ type: "text", text: recordingInfo }]
    };
  } catch (error) {
    return {
      content: [{ 
        type: "text", 
        text: `Unable to retrieve recording details. Error: ${error.message}` 
      }]
    };
  }
}

/**
 * Get meeting participants
 */
async function getMeetingParticipants(accessToken, params) {
  const { meetingId } = params;

  if (!meetingId) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: meetingId"
      }]
    };
  }

  try {
    // Convert thread ID to meeting ID if necessary
    let actualMeetingId = meetingId;

    if (meetingId.includes('@thread.v2')) {
      try {
        console.error('Attempting to convert thread ID to meeting ID for participants');
        actualMeetingId = await convertThreadIdToMeetingId(accessToken, meetingId);
        console.error('Conversion successful. New meeting ID:', actualMeetingId);
      } catch (conversionError) {
        console.error(`Unable to convert thread ID to meeting ID: ${conversionError.message}`);
        return {
          content: [{
            type: "text",
            text: `Unable to retrieve meeting participants. This might be because the meeting is still in progress or the attendance data isn't available: ${conversionError.message}`
          }]
        };
      }
    }

    // Try to get attendance records if available
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${actualMeetingId}/attendanceReports`
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ 
          type: "text", 
          text: "No attendance reports found for this meeting. Attendance reports are only available for meetings that have ended." 
        }]
      };
    }
    
    // Get the most recent report
    const latestReport = response.value[0];
    
    // Get the attendees from the report
    const attendeesResponse = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${actualMeetingId}/attendanceReports/${latestReport.id}/attendanceRecords`
    );
    
    if (!attendeesResponse.value || attendeesResponse.value.length === 0) {
      return {
        content: [{ type: "text", text: "No attendees found in the attendance report." }]
      };
    }
    
    // Format the attendees list
    const attendeesList = attendeesResponse.value.map((attendee, index) => {
      const joinTime = attendee.joinDateTime ? new Date(attendee.joinDateTime).toLocaleString() : 'Unknown';
      const leaveTime = attendee.leaveDateTime ? new Date(attendee.leaveDateTime).toLocaleString() : 'Still in meeting';
      let durationStr = 'Unknown';
      
      if (attendee.joinDateTime && attendee.leaveDateTime) {
        const joinDate = new Date(attendee.joinDateTime);
        const leaveDate = new Date(attendee.leaveDateTime);
        const durationMs = leaveDate - joinDate;
        durationStr = formatDuration(durationMs / 1000);
      }
      
      return `${index + 1}. ${attendee.emailAddress || 'Unknown'}\n   Role: ${attendee.role || 'Attendee'}\n   Joined: ${joinTime}\n   Left: ${leaveTime}\n   Duration: ${durationStr}\n`;
    }).join('\n');
    
    return {
      content: [{ 
        type: "text", 
        text: `Meeting Participants (${attendeesResponse.value.length}):\n\n${attendeesList}` 
      }]
    };
  } catch (error) {
    // This endpoint might not be available for all meetings
    return {
      content: [{ 
        type: "text", 
        text: `Unable to retrieve meeting participants. This might be because the meeting is still in progress or the attendance data isn't available: ${error.message}` 
      }]
    };
  }
}

/**
 * Get meeting insights (summary, action items, etc.)
 */
async function getMeetingInsights(accessToken, params) {
  const { meetingId } = params;
  
  if (!meetingId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: meetingId" 
      }]
    };
  }
  
  try {
    // Try to get meeting insights if available
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/onlineMeetings/${meetingId}/meetingAiInsights`
    );
    
    let insightsText = '';
    
    if (response.meetingSummary) {
      insightsText += "## Meeting Summary\n\n";
      insightsText += `${response.meetingSummary}\n\n`;
    }
    
    if (response.actionItems && response.actionItems.length > 0) {
      insightsText += "## Action Items\n\n";
      response.actionItems.forEach((item, index) => {
        insightsText += `${index + 1}. ${item.description}`;
        if (item.assignees && item.assignees.length > 0) {
          insightsText += ` (Assigned to: ${item.assignees.map(a => a.name || a.email).join(', ')})`;
        }
        insightsText += '\n';
      });
      insightsText += '\n';
    }
    
    if (response.followUps && response.followUps.length > 0) {
      insightsText += "## Follow-Ups\n\n";
      response.followUps.forEach((item, index) => {
        insightsText += `${index + 1}. ${item.description}`;
        if (item.assignees && item.assignees.length > 0) {
          insightsText += ` (For: ${item.assignees.map(a => a.name || a.email).join(', ')})`;
        }
        insightsText += '\n';
      });
      insightsText += '\n';
    }
    
    if (!insightsText) {
      return {
        content: [{ 
          type: "text", 
          text: "No meeting insights available for this meeting. Insights are typically generated after a meeting ends and might take some time to become available." 
        }]
      };
    }
    
    return {
      content: [{ 
        type: "text", 
        text: `# Meeting Insights\n\n${insightsText}` 
      }]
    };
  } catch (error) {
    return {
      content: [{ 
        type: "text", 
        text: `Unable to retrieve meeting insights. This might be because insights aren't available yet or the meeting didn't have AI insights enabled: ${error.message}` 
      }]
    };
  }
}

// Export the handler
module.exports = handleTeamsMeeting;