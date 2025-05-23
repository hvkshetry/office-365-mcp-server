/**
 * Mock data functions for test mode
 */

/**
 * Simulates Microsoft Graph API responses for testing
 * @param {string} method - HTTP method
 * @param {string} path - API path
 * @param {object} data - Request data
 * @param {object} queryParams - Query parameters
 * @returns {object} - Simulated API response
 */
function simulateGraphAPIResponse(method, path, data, queryParams) {
  console.error(`Simulating response for: ${method} ${path}`);
  
  // Email-related simulations
  if (method === 'GET') {
    if (path.includes('messages') && !path.includes('sendMail')) {
      // Simulate a successful email list/search response
      if (path.includes('/messages/')) {
        // Single email response
        return {
          id: "simulated-email-id",
          subject: "Simulated Email Subject",
          from: {
            emailAddress: {
              name: "Simulated Sender",
              address: "sender@example.com"
            }
          },
          toRecipients: [{
            emailAddress: {
              name: "Recipient Name",
              address: "recipient@example.com"
            }
          }],
          ccRecipients: [],
          bccRecipients: [],
          receivedDateTime: new Date().toISOString(),
          bodyPreview: "This is a simulated email preview...",
          body: {
            contentType: "text",
            content: "This is the full content of the simulated email. Since we can't connect to the real Microsoft Graph API, we're returning this placeholder content instead."
          },
          hasAttachments: false,
          importance: "normal",
          isRead: false,
          internetMessageHeaders: []
        };
      } else {
        // Email list response
        return {
          value: [
            {
              id: "simulated-email-1",
              subject: "Important Meeting Tomorrow",
              from: {
                emailAddress: {
                  name: "John Doe",
                  address: "john@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date().toISOString(),
              bodyPreview: "Let's discuss the project status...",
              hasAttachments: false,
              importance: "high",
              isRead: false
            },
            {
              id: "simulated-email-2",
              subject: "Weekly Report",
              from: {
                emailAddress: {
                  name: "Jane Smith",
                  address: "jane@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 86400000).toISOString(), // Yesterday
              bodyPreview: "Please find attached the weekly report...",
              hasAttachments: true,
              importance: "normal",
              isRead: true
            },
            {
              id: "simulated-email-3",
              subject: "Question about the project",
              from: {
                emailAddress: {
                  name: "Bob Johnson",
                  address: "bob@example.com"
                }
              },
              toRecipients: [{
                emailAddress: {
                  name: "You",
                  address: "you@example.com"
                }
              }],
              ccRecipients: [],
              receivedDateTime: new Date(Date.now() - 172800000).toISOString(), // 2 days ago
              bodyPreview: "I had a question about the timeline...",
              hasAttachments: false,
              importance: "normal",
              isRead: false
            }
          ]
        };
      }
    } else if (path.includes('mailFolders')) {
      // Simulate a mail folders response
      return {
        value: [
          { id: "inbox", displayName: "Inbox" },
          { id: "drafts", displayName: "Drafts" },
          { id: "sentItems", displayName: "Sent Items" },
          { id: "deleteditems", displayName: "Deleted Items" }
        ]
      };
    } else if (path.includes('/me/joinedTeams') || path.includes('/teams')) {
      // Simulate Teams response
      return {
        value: [
          {
            id: "simulated-team-1",
            displayName: "Marketing Team",
            description: "Team for marketing department",
            isArchived: false,
            visibility: "private"
          },
          {
            id: "simulated-team-2",
            displayName: "Project X",
            description: "Cross-functional team for Project X",
            isArchived: false,
            visibility: "public"
          },
          {
            id: "simulated-team-3",
            displayName: "Executive Leadership",
            description: "Leadership team",
            isArchived: false,
            visibility: "private"
          }
        ]
      };
    } else if (path.includes('channels')) {
      // Simulate channels response
      return {
        value: [
          {
            id: "simulated-channel-1",
            displayName: "General",
            description: "General channel for team discussion",
            membershipType: "standard"
          },
          {
            id: "simulated-channel-2",
            displayName: "Project Updates",
            description: "Channel for project updates and status",
            membershipType: "standard"
          },
          {
            id: "simulated-channel-3",
            displayName: "Private Discussion",
            description: "Channel for private team discussion",
            membershipType: "private"
          }
        ]
      };
    } else if (path.includes('onlineMeetings')) {
      if (path.includes('transcripts')) {
        // Simulate transcripts response
        return {
          value: [
            {
              id: "simulated-transcript-1",
              meetingId: "simulated-meeting-1",
              createdDateTime: new Date().toISOString(),
              contentCorrelationId: "correlation-id-1"
            }
          ]
        };
      } else {
        // Simulate meetings response
        return {
          value: [
            {
              id: "simulated-meeting-1",
              subject: "Weekly Status",
              startDateTime: new Date().toISOString(),
              endDateTime: new Date(Date.now() + 3600000).toISOString(),
              joinUrl: "https://teams.microsoft.com/l/meetup-join/simulated",
              participants: {
                organizer: {
                  upn: "organizer@example.com",
                  identity: {
                    displayName: "Meeting Organizer"
                  }
                },
                attendees: [
                  {
                    upn: "attendee1@example.com",
                    identity: {
                      displayName: "Attendee One"
                    }
                  },
                  {
                    upn: "attendee2@example.com",
                    identity: {
                      displayName: "Attendee Two"
                    }
                  }
                ]
              }
            }
          ]
        };
      }
    } else if (path.includes('/drive/')) {
      // Simulate OneDrive files response
      return {
        value: [
          {
            id: "simulated-file-1",
            name: "Quarterly Report.docx",
            size: 245678,
            createdDateTime: new Date(Date.now() - 8640000).toISOString(),
            lastModifiedDateTime: new Date().toISOString(),
            webUrl: "https://example.sharepoint.com/documents/quarterly-report.docx",
            file: {
              mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            }
          },
          {
            id: "simulated-folder-1",
            name: "Project Documents",
            size: 0,
            createdDateTime: new Date(Date.now() - 86400000).toISOString(),
            lastModifiedDateTime: new Date().toISOString(),
            webUrl: "https://example.sharepoint.com/documents/project-documents",
            folder: {
              childCount: 12
            }
          },
          {
            id: "simulated-file-2",
            name: "Presentation.pptx",
            size: 3456789,
            createdDateTime: new Date(Date.now() - 172800000).toISOString(),
            lastModifiedDateTime: new Date(Date.now() - 86400000).toISOString(),
            webUrl: "https://example.sharepoint.com/documents/presentation.pptx",
            file: {
              mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            }
          }
        ]
      };
    }
  } else if (method === 'POST') {
    if (path.includes('sendMail')) {
      // Simulate a successful email send
      return {};
    } else if (path.includes('/teams/') && path.includes('/channels/') && path.includes('/messages')) {
      // Simulate sending a Teams channel message
      return {
        id: "simulated-message-id",
        etag: "simulated-etag",
        messageType: "message",
        createdDateTime: new Date().toISOString(),
        lastModifiedDateTime: new Date().toISOString(),
        importance: "normal",
        subject: null,
        summary: null,
        from: {
          user: {
            id: "simulated-user-id",
            displayName: "Simulated User",
            userIdentityType: "aadUser"
          }
        },
        body: {
          contentType: "text",
          content: data?.body?.content || "Simulated message content"
        }
      };
    } else if (path.includes('/subscriptions')) {
      // Simulate creating a subscription
      return {
        id: "simulated-subscription-" + Date.now(),
        resource: data?.resource || "simulated-resource",
        changeType: data?.changeType || "created,updated",
        notificationUrl: data?.notificationUrl || "https://example.com/notifications",
        expirationDateTime: data?.expirationDateTime || new Date(Date.now() + 43200000).toISOString(), // 12 hours from now
        clientState: data?.clientState || "simulated-client-state"
      };
    }
  }
  
  // If we get here, we don't have a simulation for this endpoint
  console.error(`No simulation available for: ${method} ${path}`);
  return {};
}

module.exports = {
  simulateGraphAPIResponse
};
