/**
 * Email folder operations - move, list, create folders, folder name resolution
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { validateId } = require('../utils/validate');
const config = require('../config');

/**
 * Convert folder name to folder ID
 */
async function getFolderIdByName(accessToken, folderName, mailbox) {
  // Check well-known folder names first
  const wellKnownFolders = {
    'inbox': 'inbox',
    'sent': 'sentitems',
    'sent items': 'sentitems',
    'drafts': 'drafts',
    'deleted': 'deleteditems',
    'deleted items': 'deleteditems',
    'junk': 'junkemail',
    'junk email': 'junkemail',
    'archive': 'archive'
  };

  const lowerName = folderName.toLowerCase();
  if (wellKnownFolders[lowerName]) {
    return wellKnownFolders[lowerName];
  }

  // Search for custom folder by name
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${config.getMailboxPrefix(mailbox)}/mailFolders`,
      null,
      {
        $filter: `displayName eq '${folderName}'`,
        $select: 'id,displayName'
      }
    );

    if (response.value && response.value.length > 0) {
      return response.value[0].id;
    }
  } catch (error) {
    console.error(`Error finding folder by name: ${error.message}`);
  }

  return null;
}

/**
 * Unified email move handler with batch capability
 */
async function handleEmailMove(args) {
  const { emailIds, destinationFolderId, batch = false } = args;

  if (!emailIds || !destinationFolderId) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameters: emailIds and destinationFolderId"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    if (batch || emailIds.length > 5) {
      return await batchMoveEmails(accessToken, { emailIds, destinationFolderId });
    } else {
      return await moveEmails(accessToken, { emailIds, destinationFolderId });
    }
  } catch (error) {
    console.error('Error moving emails:', error);
    return {
      content: [{ type: "text", text: `Error moving emails: ${error.message}` }]
    };
  }
}

async function moveEmails(accessToken, params) {
  const { emailIds, destinationFolderId, mailbox } = params;

  const results = [];

  for (const emailId of emailIds) {
    validateId(emailId, 'emailId');
    try {
      await callGraphAPI(
        accessToken,
        'POST',
        `${config.getMailboxPrefix(mailbox)}/messages/${emailId}/move`,
        { destinationId: destinationFolderId }
      );
      results.push({ emailId, status: 'success' });
    } catch (error) {
      results.push({ emailId, status: 'failed', error: error.message });
    }
  }

  const successCount = results.filter(r => r.status === 'success').length;
  const failureCount = results.filter(r => r.status === 'failed').length;

  let message = `Moved ${successCount} emails successfully.`;
  if (failureCount > 0) {
    message += ` ${failureCount} emails failed to move.`;
  }

  return {
    content: [{ type: "text", text: message }]
  };
}

async function batchMoveEmails(accessToken, params) {
  const { emailIds, destinationFolderId, mailbox } = params;

  const batchSize = 20;
  const results = [];

  for (let i = 0; i < emailIds.length; i += batchSize) {
    const batch = emailIds.slice(i, i + batchSize);
    const prefix = config.getMailboxPrefix(mailbox);
    const requests = batch.map((emailId, index) => ({
      id: `${index}`,
      method: 'POST',
      url: `/${prefix}/messages/${emailId}/move`,
      body: { destinationId: destinationFolderId },
      headers: { 'Content-Type': 'application/json' }
    }));

    try {
      const batchResponse = await callGraphAPI(
        accessToken,
        'POST',
        '$batch',
        { requests }
      );

      batchResponse.responses.forEach((response, index) => {
        // Email move operations return 201 (Created) on success
        if (response.status === 200 || response.status === 201) {
          results.push({ emailId: batch[index], status: 'success' });
        } else {
          results.push({
            emailId: batch[index],
            status: 'failed',
            error: response.body?.error?.message || 'Unknown error'
          });
        }
      });
    } catch (error) {
      batch.forEach(emailId => {
        results.push({ emailId, status: 'failed', error: error.message });
      });
    }
  }

  const successCount = results.filter(r => r.status === 'success').length;
  const failureCount = results.filter(r => r.status === 'failed').length;

  let message = `Batch moved ${successCount} emails successfully.`;
  if (failureCount > 0) {
    message += ` ${failureCount} emails failed to move.`;
  }

  return {
    content: [{ type: "text", text: message }]
  };
}

/**
 * Unified email folder handler
 */
async function handleEmailFolder(args) {
  const { operation, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations are: list, create"
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'list':
        return await listEmailFolders(accessToken, params.mailbox);
      case 'create':
        return await createEmailFolder(accessToken, params);
      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid operations are: list, create`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email folder ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in email folder: ${error.message}` }]
    };
  }
}

// Folder operation implementations
async function listEmailFolders(accessToken, mailbox = null) {
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/mailFolders`,
    null,
    {
      $top: 100,
      $select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount'
    }
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No email folders found." }]
    };
  }

  const foldersList = response.value
    .filter(folder => folder.displayName !== 'Conversation History')
    .map((folder, index) => {
      return `${index + 1}. ${folder.displayName}\n   ID: ${folder.id}\n   Messages: ${folder.totalItemCount} (${folder.unreadItemCount} unread)\n`;
    }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} email folders:\n\n${foldersList}`
    }]
  };
}

async function createEmailFolder(accessToken, params) {
  const { displayName, parentFolderId, mailbox } = params;

  if (!displayName) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: displayName"
      }]
    };
  }

  if (parentFolderId) validateId(parentFolderId, 'parentFolderId');
  const endpoint = parentFolderId
    ? `${config.getMailboxPrefix(mailbox)}/mailFolders/${parentFolderId}/childFolders`
    : `${config.getMailboxPrefix(mailbox)}/mailFolders`;

  const folderData = { displayName };

  const response = await callGraphAPI(
    accessToken,
    'POST',
    endpoint,
    folderData
  );

  return {
    content: [{
      type: "text",
      text: `Email folder created successfully!\nFolder ID: ${response.id}`
    }]
  };
}

module.exports = {
  getFolderIdByName,
  handleEmailMove,
  handleEmailFolder
};
