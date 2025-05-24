/**
 * Consolidated Teams Channel Module
 * 
 * Provides a unified interface for all Teams channel operations including:
 * - Listing channels in a team
 * - Creating new channels
 * - Getting channel details
 * - Updating channels
 * - Deleting channels
 * - Working with channel messages
 * - Managing channel members
 * - Accessing channel tabs
 */

const { ensureAuthenticated } = require('../../auth');
const { callGraphAPI } = require('../../utils/graph-api');
const config = require('../../config');

/**
 * Main handler for teams_channel operations
 * @param {Object} args - The operation arguments
 * @returns {Object} - MCP response
 */
async function handleTeamsChannel(args) {
  const { operation, ...params } = args;
  
  if (!operation) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: operation. Valid operations are: list, create, get, update, delete, list_messages, get_message, create_message, reply_to_message, list_members, add_member, remove_member, list_tabs" 
      }]
    };
  }
  
  try {
    console.error(`teams_channel operation: ${operation}`);
    console.error('teams_channel params:', JSON.stringify(params));
    
    const accessToken = await ensureAuthenticated();
    
    switch (operation) {
      case 'list':
        return await listChannels(accessToken, params);
      case 'create':
        return await createChannel(accessToken, params);
      case 'get':
        return await getChannel(accessToken, params);
      case 'update':
        return await updateChannel(accessToken, params);
      case 'delete':
        return await deleteChannel(accessToken, params);
      case 'list_messages':
        return await listChannelMessages(accessToken, params);
      case 'get_message':
        return await getChannelMessage(accessToken, params);
      case 'create_message':
        return await createChannelMessage(accessToken, params);
      case 'reply_to_message':
        return await replyToChannelMessage(accessToken, params);
      case 'list_members':
        return await listChannelMembers(accessToken, params);
      case 'add_member':
        return await addChannelMember(accessToken, params);
      case 'remove_member':
        return await removeChannelMember(accessToken, params);
      case 'list_tabs':
        return await listChannelTabs(accessToken, params);
      default:
        return {
          content: [{ 
            type: "text", 
            text: `Invalid operation: ${operation}. Valid operations are: list, create, get, update, delete, list_messages, get_message, create_message, reply_to_message, list_members, add_member, remove_member, list_tabs` 
          }]
        };
    }
  } catch (error) {
    console.error(`Error in teams_channel ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in teams_channel operation: ${error.message}` }]
    };
  }
}

/**
 * List channels in a team
 */
async function listChannels(accessToken, params) {
  const { teamId } = params;
  
  if (!teamId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameter: teamId" 
      }]
    };
  }
  
  // The channels endpoint doesn't support $top query parameter
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `teams/${teamId}/channels`,
    null,
    {
      $select: 'id,displayName,description,email,webUrl,membershipType'
    }
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No channels found in this team." }]
    };
  }
  
  // Format the channel list
  const channelList = response.value.map((channel, index) => {
    return `${index + 1}. ${channel.displayName}\n   ID: ${channel.id}\n   Type: ${channel.membershipType || 'standard'}\n   Description: ${channel.description || 'No description'}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} channels in team:\n\n${channelList}` 
    }]
  };
}

/**
 * Create a new channel
 */
async function createChannel(accessToken, params) {
  const { teamId, displayName, description, membershipType = 'standard' } = params;
  
  if (!teamId || !displayName) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and displayName." 
      }]
    };
  }
  
  // Only standard and private membership types are supported
  if (membershipType !== 'standard' && membershipType !== 'private') {
    return {
      content: [{ 
        type: "text", 
        text: "Invalid membershipType. Must be 'standard' or 'private'." 
      }]
    };
  }
  
  const newChannel = {
    displayName,
    membershipType
  };
  
  if (description) {
    newChannel.description = description;
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `teams/${teamId}/channels`,
    newChannel
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Channel created successfully!\nName: ${response.displayName}\nID: ${response.id}\nType: ${response.membershipType}` 
    }]
  };
}

/**
 * Get channel details
 */
async function getChannel(accessToken, params) {
  const { teamId, channelId } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  const channel = await callGraphAPI(
    accessToken,
    'GET',
    `teams/${teamId}/channels/${channelId}`
  );
  
  let channelInfo = `Channel: ${channel.displayName}\n`;
  channelInfo += `ID: ${channel.id}\n`;
  channelInfo += `Type: ${channel.membershipType || 'standard'}\n`;
  channelInfo += `Description: ${channel.description || 'No description'}\n`;
  
  if (channel.email) {
    channelInfo += `Email: ${channel.email}\n`;
  }
  
  if (channel.webUrl) {
    channelInfo += `URL: ${channel.webUrl}\n`;
  }

  return {
    content: [{ type: "text", text: channelInfo }]
  };
}

/**
 * Update a channel
 */
async function updateChannel(accessToken, params) {
  const { teamId, channelId, displayName, description } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  const updateData = {};
  if (displayName) updateData.displayName = displayName;
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
    `teams/${teamId}/channels/${channelId}`,
    updateData
  );
  
  return {
    content: [{ type: "text", text: "Channel updated successfully!" }]
  };
}

/**
 * Delete a channel
 */
async function deleteChannel(accessToken, params) {
  const { teamId, channelId } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  await callGraphAPI(
    accessToken,
    'DELETE',
    `teams/${teamId}/channels/${channelId}`
  );
  
  return {
    content: [{ type: "text", text: "Channel deleted successfully!" }]
  };
}

/**
 * List messages in a channel
 */
async function listChannelMessages(accessToken, params) {
  const { teamId, channelId, maxResults = 25 } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `teams/${teamId}/channels/${channelId}/messages`,
    null,
    {
      $top: maxResults
    }
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No messages found in this channel." }]
    };
  }
  
  // Format the message list
  const messageList = response.value.map((message, index) => {
    const sender = message.from?.user?.displayName || message.from?.user?.id || message.from?.application?.displayName || 'Unknown';
    const createdTime = new Date(message.createdDateTime).toLocaleString();
    const replyCount = message.replies?.length || 0;
    const contentPreview = message.body?.content ? truncateContent(message.body.content, 100) : 'No content';
    
    return `${index + 1}. From: ${sender} (${createdTime})\n   ID: ${message.id}\n   ${replyCount > 0 ? `Replies: ${replyCount}` : 'No replies'}\n   ${contentPreview}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} messages in channel:\n\n${messageList}` 
    }]
  };
}

/**
 * Get a specific channel message
 */
async function getChannelMessage(accessToken, params) {
  const { teamId, channelId, messageId } = params;
  
  if (!teamId || !channelId || !messageId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId, channelId, and messageId." 
      }]
    };
  }
  
  const message = await callGraphAPI(
    accessToken,
    'GET',
    `teams/${teamId}/channels/${channelId}/messages/${messageId}`
  );
  
  const sender = message.from?.user?.displayName || message.from?.user?.id || message.from?.application?.displayName || 'Unknown';
  const createdTime = new Date(message.createdDateTime).toLocaleString();
  
  let content = 'No content';
  if (message.body?.content) {
    // Attempt to remove HTML from the content
    content = message.body.content
      .replace(/<[^>]*>/g, ' ') // Replace HTML tags with space
      .replace(/&nbsp;/g, ' ')  // Replace &nbsp; with space
      .replace(/\s+/g, ' ')     // Collapse multiple spaces
      .trim();
  }
  
  // Check for attachments
  let attachments = 'No attachments';
  if (message.attachments && message.attachments.length > 0) {
    attachments = message.attachments.map((attachment, index) => {
      return `   ${index + 1}. ${attachment.name || 'Unnamed'} (${attachment.contentType || 'Unknown type'})`;
    }).join('\n');
  }
  
  // Get replies if available
  let replies;
  try {
    const repliesResponse = await callGraphAPI(
      accessToken,
      'GET',
      `teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`
    );
    
    if (repliesResponse.value && repliesResponse.value.length > 0) {
      replies = repliesResponse.value.map((reply, index) => {
        const replySender = reply.from?.user?.displayName || reply.from?.user?.id || reply.from?.application?.displayName || 'Unknown';
        const replyTime = new Date(reply.createdDateTime).toLocaleString();
        const replyContent = reply.body?.content ? truncateContent(reply.body.content.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim(), 100) : 'No content';
        
        return `   ${index + 1}. From: ${replySender} (${replyTime})\n      ${replyContent}`;
      }).join('\n\n');
    } else {
      replies = 'No replies';
    }
  } catch (error) {
    console.error('Error fetching replies:', error);
    replies = 'Unable to fetch replies';
  }
  
  const messageDetails = `From: ${sender}\nTime: ${createdTime}\nID: ${message.id}\n\nContent:\n${content}\n\nAttachments:\n${attachments}\n\nReplies:\n${replies}`;
  
  return {
    content: [{ type: "text", text: messageDetails }]
  };
}

/**
 * Create a new channel message
 */
async function createChannelMessage(accessToken, params) {
  const { teamId, channelId, content, attachments } = params;
  
  if (!teamId || !channelId || !content) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId, channelId, and content." 
      }]
    };
  }
  
  const messageData = {
    body: {
      content,
      contentType: content.includes('<') ? 'html' : 'text'
    }
  };
  
  // Add attachments if present
  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    messageData.attachments = attachments.map(attachment => {
      const { name, contentUrl, contentType } = attachment;
      return { name, contentUrl, contentType };
    });
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `teams/${teamId}/channels/${channelId}/messages`,
    messageData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Message posted successfully!\nID: ${response.id}` 
    }]
  };
}

/**
 * Reply to a channel message
 */
async function replyToChannelMessage(accessToken, params) {
  const { teamId, channelId, messageId, content, attachments } = params;
  
  if (!teamId || !channelId || !messageId || !content) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId, channelId, messageId, and content." 
      }]
    };
  }
  
  const replyData = {
    body: {
      content,
      contentType: content.includes('<') ? 'html' : 'text'
    }
  };
  
  // Add attachments if present
  if (attachments && Array.isArray(attachments) && attachments.length > 0) {
    replyData.attachments = attachments.map(attachment => {
      const { name, contentUrl, contentType } = attachment;
      return { name, contentUrl, contentType };
    });
  }
  
  const response = await callGraphAPI(
    accessToken,
    'POST',
    `teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`,
    replyData
  );
  
  return {
    content: [{ 
      type: "text", 
      text: `Reply posted successfully!\nID: ${response.id}` 
    }]
  };
}

/**
 * List channel members
 */
async function listChannelMembers(accessToken, params) {
  const { teamId, channelId } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  // Note: This endpoint only works for private channels (membershipType=private)
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `teams/${teamId}/channels/${channelId}/members`
    );
    
    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No members found in this channel." }]
      };
    }
    
    // Format the member list
    const memberList = response.value.map((member, index) => {
      const displayName = member.displayName || 'Unknown';
      const email = member.email || 'No email';
      const userId = member.userId || 'No ID';
      const roles = member.roles?.length > 0 ? member.roles.join(', ') : 'Member';
      
      return `${index + 1}. ${displayName}\n   Email: ${email}\n   ID: ${userId}\n   Roles: ${roles}\n`;
    }).join('\n');
    
    return {
      content: [{ 
        type: "text", 
        text: `Found ${response.value.length} members in channel:\n\n${memberList}` 
      }]
    };
  } catch (error) {
    // If the channel is standard, this endpoint will return an error
    if (error.message.includes('not private')) {
      return {
        content: [{ type: "text", text: "Member listing is only available for private channels. This appears to be a standard channel." }]
      };
    }
    throw error;
  }
}

/**
 * Add a member to a private channel
 */
async function addChannelMember(accessToken, params) {
  const { teamId, channelId, userId, email, displayName, roles = [] } = params;
  
  if (!teamId || !channelId || (!userId && !email)) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId, channelId, and either userId or email." 
      }]
    };
  }
  
  // Determine the member to add
  let userIdentifier;
  if (userId) {
    userIdentifier = userId;
  } else if (email) {
    // Try to resolve email to user ID
    try {
      const userResponse = await callGraphAPI(
        accessToken,
        'GET',
        'users',
        null,
        {
          $filter: `mail eq '${email}' or userPrincipalName eq '${email}'`,
          $select: 'id'
        }
      );

      if (userResponse.value && userResponse.value.length > 0) {
        userIdentifier = userResponse.value[0].id;
      } else {
        return {
          content: [{ type: 'text', text: `Unable to find a user with email: ${email}` }]
        };
      }
    } catch (error) {
      console.error('Error resolving email to user ID:', error);
      return {
        content: [{ type: 'text', text: `Error resolving email to user ID: ${error.message}` }]
      };
    }
  }

  const memberData = {
    '@odata.type': '#microsoft.graph.aadUserConversationMember',
    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userIdentifier}')`
  };

  if (displayName) memberData.displayName = displayName;
  if (roles && Array.isArray(roles) && roles.length > 0) memberData.roles = roles;
  
  try {
    const response = await callGraphAPI(
      accessToken,
      'POST',
      `teams/${teamId}/channels/${channelId}/members`,
      memberData
    );
    
    return {
      content: [{ 
        type: "text", 
        text: `Member added successfully!\nID: ${response.id}` 
      }]
    };
  } catch (error) {
    // If the channel is standard, this endpoint will return an error
    if (error.message.includes('not private')) {
      return {
        content: [{ type: "text", text: "Member management is only available for private channels. This appears to be a standard channel." }]
      };
    }
    throw error;
  }
}

/**
 * Remove a member from a private channel
 */
async function removeChannelMember(accessToken, params) {
  const { teamId, channelId, memberId } = params;
  
  if (!teamId || !channelId || !memberId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId, channelId, and memberId." 
      }]
    };
  }
  
  try {
    await callGraphAPI(
      accessToken,
      'DELETE',
      `teams/${teamId}/channels/${channelId}/members/${memberId}`
    );
    
    return {
      content: [{ type: "text", text: "Member removed successfully!" }]
    };
  } catch (error) {
    // If the channel is standard, this endpoint will return an error
    if (error.message.includes('not private')) {
      return {
        content: [{ type: "text", text: "Member management is only available for private channels. This appears to be a standard channel." }]
      };
    }
    throw error;
  }
}

/**
 * List channel tabs
 */
async function listChannelTabs(accessToken, params) {
  const { teamId, channelId } = params;
  
  if (!teamId || !channelId) {
    return {
      content: [{ 
        type: "text", 
        text: "Missing required parameters. Please provide teamId and channelId." 
      }]
    };
  }
  
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `teams/${teamId}/channels/${channelId}/tabs`
  );
  
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No tabs found in this channel." }]
    };
  }
  
  // Format the tab list
  const tabList = response.value.map((tab, index) => {
    return `${index + 1}. ${tab.displayName || 'Unnamed tab'}\n   ID: ${tab.id}\n   Type: ${tab.teamsAppId || 'Unknown'}\n   URL: ${tab.webUrl || 'No URL'}\n`;
  }).join('\n');
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} tabs in channel:\n\n${tabList}` 
    }]
  };
}

/**
 * Helper function to truncate content to a specific length
 */
function truncateContent(content, maxLength) {
  if (!content) return 'No content';
  
  // Remove HTML tags
  const plainText = content
    .replace(/<[^>]*>/g, ' ')  // Replace HTML tags with space
    .replace(/&nbsp;/g, ' ')   // Replace &nbsp; with space
    .replace(/\s+/g, ' ')      // Collapse multiple spaces
    .trim();
  
  if (plainText.length <= maxLength) {
    return plainText;
  }
  
  return plainText.substring(0, maxLength) + '...';
}

// Export the handler
module.exports = handleTeamsChannel;