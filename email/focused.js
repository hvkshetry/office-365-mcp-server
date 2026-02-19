/**
 * Focused inbox retrieval
 */

const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');

/**
 * Get focused inbox messages
 */
async function getFocusedInbox(accessToken, params) {
  const { maxResults = 25, mailbox } = params;

  const queryParams = {
    $filter: 'inferenceClassification eq \'Focused\'',
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime DESC',
    $top: maxResults
  };

  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/messages`,
    null,
    queryParams
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No focused messages found." }]
    };
  }

  const messagesList = response.value.map((msg, index) => {
    return `${index + 1}. ${msg.subject}
   From: ${msg.from?.emailAddress?.address || 'N/A'}
   Date: ${new Date(msg.receivedDateTime).toLocaleString()}
   Preview: ${msg.bodyPreview?.substring(0, 100)}...
   Has Attachments: ${msg.hasAttachments ? 'Yes' : 'No'}
   ID: ${msg.id}`;
  }).join('\n\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} focused messages:\n\n${messagesList}`
    }]
  };
}

module.exports = { getFocusedInbox };
