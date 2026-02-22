/**
 * Microsoft 365 Groups module
 * Provides access to group management (foundation for Teams/Planner)
 * Graph API: /groups, /members, /owners
 */
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { safeTool } = require('../utils/errors');

async function handleGroups(args) {
  const { operation, ...params } = args || {};

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: list, get, create, update, delete, list_members, add_member, remove_member, list_owners, get_drive, list_drives"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'list': {
        const queryParams = {
          $top: params.maxResults || 25,
          $select: 'id,displayName,description,mail,groupTypes,visibility,createdDateTime'
        };

        if (params.filter) {
          queryParams.$filter = params.filter;
        } else if (params.displayName) {
          queryParams.$filter = `startswith(displayName, '${params.displayName}')`;
        }

        // By default, list only Microsoft 365 groups (unified groups)
        if (!params.filter && !params.includeAll) {
          queryParams.$filter = queryParams.$filter
            ? `${queryParams.$filter} and groupTypes/any(g:g eq 'Unified')`
            : "groupTypes/any(g:g eq 'Unified')";
        }

        const response = await callGraphAPI(accessToken, 'GET', 'groups', null, queryParams);

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No groups found." }] };
        }

        const list = response.value.map(g => {
          const visibility = g.visibility ? ` (${g.visibility})` : '';
          return `- ${g.displayName}${visibility}\n  Mail: ${g.mail || 'N/A'}\n  ID: ${g.id}`;
        }).join('\n');

        return { content: [{ type: "text", text: `Found ${response.value.length} groups:\n\n${list}` }] };
      }

      case 'get': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        const group = await callGraphAPI(accessToken, 'GET', `groups/${params.groupId}`, null, {
          $select: 'id,displayName,description,mail,groupTypes,visibility,createdDateTime,mailNickname'
        });

        const info = [
          `Name: ${group.displayName}`,
          `Description: ${group.description || 'None'}`,
          `Mail: ${group.mail || 'N/A'}`,
          `Visibility: ${group.visibility || 'N/A'}`,
          `Created: ${new Date(group.createdDateTime).toLocaleString()}`,
          `ID: ${group.id}`
        ].join('\n');

        return { content: [{ type: "text", text: `Group Details:\n\n${info}` }] };
      }

      case 'create': {
        if (!params.displayName || !params.mailNickname) {
          return { content: [{ type: "text", text: "Missing required parameters: displayName, mailNickname" }] };
        }

        const groupData = {
          displayName: params.displayName,
          mailNickname: params.mailNickname,
          mailEnabled: true,
          securityEnabled: false,
          groupTypes: ['Unified'],
          visibility: params.visibility || 'Private'
        };

        if (params.description) groupData.description = params.description;
        if (params.owners && params.owners.length > 0) {
          groupData['owners@odata.bind'] = params.owners.map(id =>
            `https://graph.microsoft.com/v1.0/users/${id}`
          );
        }
        if (params.members && params.members.length > 0) {
          groupData['members@odata.bind'] = params.members.map(id =>
            `https://graph.microsoft.com/v1.0/users/${id}`
          );
        }

        const newGroup = await callGraphAPI(accessToken, 'POST', 'groups', groupData);

        return { content: [{ type: "text", text: `Group created!\nID: ${newGroup.id}\nName: ${newGroup.displayName}` }] };
      }

      case 'update': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        const updateData = {};
        if (params.displayName) updateData.displayName = params.displayName;
        if (params.description !== undefined) updateData.description = params.description;
        if (params.visibility) updateData.visibility = params.visibility;

        await callGraphAPI(accessToken, 'PATCH', `groups/${params.groupId}`, updateData);

        return { content: [{ type: "text", text: "Group updated successfully!" }] };
      }

      case 'delete': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        await callGraphAPI(accessToken, 'DELETE', `groups/${params.groupId}`, null);

        return { content: [{ type: "text", text: "Group deleted successfully!" }] };
      }

      case 'list_members': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        const response = await callGraphAPI(accessToken, 'GET',
          `groups/${params.groupId}/members`, null, {
            $top: params.maxResults || 100,
            $select: 'id,displayName,mail,userPrincipalName'
          }
        );

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No members found." }] };
        }

        const memberList = response.value.map(m =>
          `- ${m.displayName} <${m.mail || m.userPrincipalName || 'N/A'}> (ID: ${m.id})`
        ).join('\n');

        return { content: [{ type: "text", text: `${response.value.length} members:\n\n${memberList}` }] };
      }

      case 'add_member': {
        if (!params.groupId || !params.userId) {
          return { content: [{ type: "text", text: "Missing required parameters: groupId, userId" }] };
        }

        await callGraphAPI(accessToken, 'POST',
          `groups/${params.groupId}/members/$ref`,
          { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${params.userId}` }
        );

        return { content: [{ type: "text", text: "Member added successfully!" }] };
      }

      case 'remove_member': {
        if (!params.groupId || !params.userId) {
          return { content: [{ type: "text", text: "Missing required parameters: groupId, userId" }] };
        }

        await callGraphAPI(accessToken, 'DELETE',
          `groups/${params.groupId}/members/${params.userId}/$ref`, null
        );

        return { content: [{ type: "text", text: "Member removed successfully!" }] };
      }

      case 'get_drive': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        try {
          const drive = await callGraphAPI(accessToken, 'GET',
            `groups/${params.groupId}/drive`, null, {
              $select: 'id,name,driveType,webUrl,owner,quota'
            });

          // Also fetch the site ID for easier chaining
          let siteInfo = '';
          try {
            const site = await callGraphAPI(accessToken, 'GET',
              `groups/${params.groupId}/sites/root`, null, {
                $select: 'id,displayName,webUrl'
              });
            siteInfo = `\nSite ID: ${site.id}\nSite Name: ${site.displayName}\nSite URL: ${site.webUrl}`;
          } catch (e) {
            // Site lookup may fail for some group types — non-fatal
          }

          const info = [
            `Drive: ${drive.name}`,
            `Drive ID: ${drive.id}`,
            `Type: ${drive.driveType}`,
            `Web URL: ${drive.webUrl}`
          ].join('\n');

          return { content: [{ type: "text", text: `${info}${siteInfo}` }] };
        } catch (error) {
          if (error.message?.includes('404') || error.statusCode === 404) {
            return { content: [{ type: "text", text: "This group does not have a SharePoint drive." }] };
          }
          throw error;
        }
      }

      case 'list_drives': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        try {
          const response = await callGraphAPI(accessToken, 'GET',
            `groups/${params.groupId}/drives`, null, {
              $select: 'id,name,driveType,webUrl'
            });

          if (!response.value || response.value.length === 0) {
            return { content: [{ type: "text", text: "No drives found for this group." }] };
          }

          const driveList = response.value.map(d =>
            `- ${d.name}\n  Drive ID: ${d.id}\n  Type: ${d.driveType}\n  URL: ${d.webUrl}`
          ).join('\n');

          return { content: [{ type: "text", text: `${response.value.length} drives:\n\n${driveList}` }] };
        } catch (error) {
          if (error.message?.includes('404') || error.statusCode === 404) {
            return { content: [{ type: "text", text: "This group does not have any SharePoint drives." }] };
          }
          throw error;
        }
      }

      case 'list_owners': {
        if (!params.groupId) {
          return { content: [{ type: "text", text: "Missing required parameter: groupId" }] };
        }

        const response = await callGraphAPI(accessToken, 'GET',
          `groups/${params.groupId}/owners`, null, {
            $select: 'id,displayName,mail,userPrincipalName'
          }
        );

        if (!response.value || response.value.length === 0) {
          return { content: [{ type: "text", text: "No owners found." }] };
        }

        const ownerList = response.value.map(o =>
          `- ${o.displayName} <${o.mail || o.userPrincipalName || 'N/A'}> (ID: ${o.id})`
        ).join('\n');

        return { content: [{ type: "text", text: `${response.value.length} owners:\n\n${ownerList}` }] };
      }

      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid: list, get, create, update, delete, list_members, add_member, remove_member, list_owners, get_drive, list_drives`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in groups ${operation}:`, error);
    return { content: [{ type: "text", text: `Error in groups ${operation}: ${error.message}` }] };
  }
}

const groupsTools = [
  {
    name: 'groups',
    description: 'Microsoft 365 Groups: list, create, manage members and owners, resolve SharePoint drives',
    inputSchema: {
      type: 'object',
      properties: {
        operation: {
          type: 'string',
          enum: ['list', 'get', 'create', 'update', 'delete', 'list_members', 'add_member', 'remove_member', 'list_owners', 'get_drive', 'list_drives'],
          description: 'Operation to perform'
        },
        groupId: { type: 'string', description: 'Group ID' },
        displayName: { type: 'string', description: 'Group display name' },
        description: { type: 'string', description: 'Group description' },
        mailNickname: { type: 'string', description: 'Mail nickname (for create, no spaces)' },
        visibility: { type: 'string', enum: ['Private', 'Public'], description: 'Group visibility' },
        userId: { type: 'string', description: 'User ID (for add/remove member)' },
        owners: { type: 'array', items: { type: 'string' }, description: 'Owner user IDs (for create)' },
        members: { type: 'array', items: { type: 'string' }, description: 'Member user IDs (for create)' },
        filter: { type: 'string', description: 'OData filter expression' },
        includeAll: { type: 'boolean', description: 'Include all group types (not just M365)' },
        maxResults: { type: 'number', description: 'Max results (default: 25)' }
      },
      required: ['operation']
    },
    handler: safeTool('groups', handleGroups)
  }
];

module.exports = { groupsTools };
