/**
 * Directory module - User profiles, presence, org hierarchy
 * Graph API: /users, /presence, /photo, /manager, /directReports
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { safeTool } = require('../utils/errors');

async function handleDirectory(args) {
  const { operation, ...params } = args || {};

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: lookup_user, get_profile, get_manager, get_reports, get_presence, search_users"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'lookup_user': {
        const identifier = params.email || params.userId;
        if (!identifier) {
          return { content: [{ type: "text", text: "Missing required parameter: email or userId" }] };
        }

        const user = await callGraphAPI(accessToken, 'GET', `users/${identifier}`, null, {
          $select: 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,companyName'
        });

        const info = [
          `Name: ${user.displayName}`,
          `Email: ${user.mail || user.userPrincipalName}`,
          user.jobTitle ? `Title: ${user.jobTitle}` : null,
          user.department ? `Department: ${user.department}` : null,
          user.companyName ? `Company: ${user.companyName}` : null,
          user.officeLocation ? `Office: ${user.officeLocation}` : null,
          user.mobilePhone ? `Mobile: ${user.mobilePhone}` : null,
          user.businessPhones?.length > 0 ? `Phone: ${user.businessPhones.join(', ')}` : null,
          `ID: ${user.id}`
        ].filter(Boolean).join('\n');

        return { content: [{ type: "text", text: `User Profile:\n\n${info}` }] };
      }

      case 'get_profile': {
        // Get current user's profile (or specified user)
        const endpoint = params.userId ? `users/${params.userId}` : 'me';

        const user = await callGraphAPI(accessToken, 'GET', endpoint, null, {
          $select: 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,companyName,aboutMe,birthday,hireDate,interests,skills,schools'
        });

        const info = [
          `Name: ${user.displayName}`,
          `Email: ${user.mail || user.userPrincipalName}`,
          user.jobTitle ? `Title: ${user.jobTitle}` : null,
          user.department ? `Department: ${user.department}` : null,
          user.companyName ? `Company: ${user.companyName}` : null,
          user.officeLocation ? `Office: ${user.officeLocation}` : null,
          user.aboutMe ? `About: ${user.aboutMe}` : null,
          user.hireDate ? `Hire Date: ${user.hireDate}` : null,
          user.skills?.length > 0 ? `Skills: ${user.skills.join(', ')}` : null,
          user.interests?.length > 0 ? `Interests: ${user.interests.join(', ')}` : null,
          `ID: ${user.id}`
        ].filter(Boolean).join('\n');

        return { content: [{ type: "text", text: `User Profile:\n\n${info}` }] };
      }

      case 'get_manager': {
        const userRef = params.userId || params.email || 'me';
        const endpoint = userRef === 'me' ? 'me/manager' : `users/${userRef}/manager`;

        try {
          const manager = await callGraphAPI(accessToken, 'GET', endpoint, null, {
            $select: 'id,displayName,mail,jobTitle,department'
          });

          return {
            content: [{
              type: "text",
              text: `Manager:\n${manager.displayName} <${manager.mail}>\nTitle: ${manager.jobTitle || 'N/A'}\nDepartment: ${manager.department || 'N/A'}\nID: ${manager.id}`
            }]
          };
        } catch (error) {
          if (error.message.includes('404')) {
            return { content: [{ type: "text", text: "No manager found for this user." }] };
          }
          throw error;
        }
      }

      case 'get_reports': {
        const userRef = params.userId || params.email || 'me';
        const endpoint = userRef === 'me' ? 'me/directReports' : `users/${userRef}/directReports`;

        const response = await callGraphAPI(accessToken, 'GET', endpoint, null, {
          $select: 'id,displayName,mail,jobTitle,department'
        });

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No direct reports found." }] };
        }

        const reports = response.value.map(r =>
          `- ${r.displayName} <${r.mail || 'N/A'}>\n  Title: ${r.jobTitle || 'N/A'} | Dept: ${r.department || 'N/A'}\n  ID: ${r.id}`
        ).join('\n');

        return { content: [{ type: "text", text: `${response.value.length} direct reports:\n\n${reports}` }] };
      }

      case 'get_presence': {
        const userRef = params.userId || 'me';

        let presence;
        if (userRef === 'me') {
          presence = await callGraphAPI(accessToken, 'GET', 'me/presence', null);
        } else {
          // For other users, use communications/presences endpoint
          const response = await callGraphAPI(accessToken, 'POST',
            'communications/getPresencesByUserId',
            { ids: Array.isArray(userRef) ? userRef : [userRef] }
          );
          presence = response.value ? response.value[0] : response;
        }

        const statusEmoji = {
          'Available': 'Available',
          'Busy': 'Busy',
          'DoNotDisturb': 'Do Not Disturb',
          'BeRightBack': 'Be Right Back',
          'Away': 'Away',
          'Offline': 'Offline'
        };

        const status = statusEmoji[presence.availability] || presence.availability;
        const activity = presence.activity || 'Unknown';

        return {
          content: [{
            type: "text",
            text: `Presence: ${status}\nActivity: ${activity}`
          }]
        };
      }

      case 'search_users': {
        if (!params.query) {
          return { content: [{ type: "text", text: "Missing required parameter: query" }] };
        }

        const queryParams = {
          $top: params.maxResults || 25,
          $select: 'id,displayName,mail,userPrincipalName,jobTitle,department',
          $filter: `startswith(displayName, '${params.query}') or startswith(mail, '${params.query}') or startswith(userPrincipalName, '${params.query}')`
        };

        const response = await callGraphAPI(accessToken, 'GET', 'users', null, queryParams);

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No users found matching your query." }] };
        }

        const userList = response.value.map(u =>
          `- ${u.displayName} <${u.mail || u.userPrincipalName}>\n  Title: ${u.jobTitle || 'N/A'} | Dept: ${u.department || 'N/A'}\n  ID: ${u.id}`
        ).join('\n');

        return { content: [{ type: "text", text: `Found ${response.value.length} users:\n\n${userList}` }] };
      }

      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid: lookup_user, get_profile, get_manager, get_reports, get_presence, search_users`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in directory ${operation}:`, error);
    return { content: [{ type: "text", text: `Error in directory ${operation}: ${error.message}` }] };
  }
}

const directoryTools = [
  {
    name: 'directory',
    description: 'User directory: profiles, managers, direct reports, presence, and user search',
    inputSchema: {
      type: 'object',
      properties: {
        operation: {
          type: 'string',
          enum: ['lookup_user', 'get_profile', 'get_manager', 'get_reports', 'get_presence', 'search_users'],
          description: 'Operation to perform'
        },
        email: { type: 'string', description: 'User email address' },
        userId: { type: 'string', description: 'User ID (or "me" for current user)' },
        query: { type: 'string', description: 'Search query (for search_users)' },
        maxResults: { type: 'number', description: 'Max results (default: 25)' }
      },
      required: ['operation']
    },
    handler: safeTool('directory', handleDirectory)
  }
];

module.exports = { directoryTools };
