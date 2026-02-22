/**
 * Consolidated Email module - thin router
 * Delegates to submodules: search, folders, rules, categories, attachments, focused
 */

const fs = require('fs');
const path = require('path');
const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { safeTool } = require('../utils/errors');
const config = require('../config');

// Submodule imports
const { handleEmailSearch } = require('./search');
const { handleEmailMove, handleEmailFolder } = require('./folders');
const { handleEmailRules } = require('./rules');
const { handleEmailCategories } = require('./categories');
const { getFocusedInbox } = require('./focused');
const { convertSharePointUrlToLocal, downloadEmbeddedAttachment, cleanupOldAttachments } = require('./attachments');

// ============== CORE EMAIL CRUD ==============

/**
 * Core email handler for list, read, send, reply, and draft operations
 */
async function handleEmail(args) {
  if (!args || typeof args !== 'object') {
    return {
      content: [{
        type: "text",
        text: "Invalid args: expected an object with 'operation' parameter"
      }]
    };
  }

  const { operation, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations are: list, read, get_attachment, cleanup_attachment, send, reply, draft, update_draft, send_draft, list_drafts"
      }]
    };
  }

  console.error(`[EMAIL] Operation: ${operation}`);

  try {
    const accessToken = await ensureAuthenticated();

    switch (operation) {
      case 'list':
        return await listEmails(accessToken, params);
      case 'read':
        return await readEmail(accessToken, params);
      case 'get_attachment':
        return await getAttachment(accessToken, params);
      case 'cleanup_attachment':
        return cleanupAttachment(params);
      case 'send':
        console.error('[EMAIL] Sending email');
        return await sendEmail(accessToken, params);
      case 'reply':
        console.error('[EMAIL] Replying to email');
        return await replyToEmail(accessToken, params);
      case 'draft':
        console.error('[EMAIL] Creating draft');
        return await createDraft(accessToken, params);
      case 'update_draft':
        return await updateDraft(accessToken, params);
      case 'send_draft':
        return await sendDraft(accessToken, params);
      case 'list_drafts':
        return await listDrafts(accessToken, params);
      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid operations are: list, read, send, reply, draft, update_draft, send_draft, list_drafts`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in email ${operation}:`, error);
    console.error('Error stack:', error.stack);
    return {
      content: [{ type: "text", text: `Error in email ${operation}: ${error.message}` }]
    };
  }
}

async function listEmails(accessToken, params) {
  const { folderId, maxResults = 10, mailbox } = params;

  const endpoint = folderId ?
    `${config.getMailboxPrefix(mailbox)}/mailFolders/${folderId}/messages` :
    `${config.getMailboxPrefix(mailbox)}/messages`;

  const queryParams = {
    $top: maxResults,
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };

  const response = await callGraphAPI(
    accessToken,
    'GET',
    endpoint,
    null,
    queryParams
  );

  if (!response.value || response.value.length === 0) {
    return {
      content: [{ type: "text", text: "No emails found." }]
    };
  }

  const emailsList = response.value.map(email => {
    const attachments = email.hasAttachments ? ' 📎' : '';
    const fromAddress = email.from?.emailAddress?.address || email.from?.address || 'Unknown sender';
    const fromName = email.from?.emailAddress?.name || email.from?.name || '';
    const fromDisplay = fromName ? `${fromName} <${fromAddress}>` : fromAddress;
    return `- ${email.subject || '(No subject)'}${attachments}\n  From: ${fromDisplay}\n  Date: ${new Date(email.receivedDateTime).toLocaleString()}\n  ID: ${email.id}\n`;
  }).join('\n');

  return {
    content: [{
      type: "text",
      text: `Found ${response.value.length} emails:\n\n${emailsList}`
    }]
  };
}

async function readEmail(accessToken, params) {
  const { emailId, mailbox } = params;

  // Clean up old attachments periodically
  try {
    cleanupOldAttachments();
  } catch (err) {
    console.error('Error during attachment cleanup:', err);
  }

  if (!emailId) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: emailId"
      }]
    };
  }

  // Fetch metadata only — attachments listed without downloading content
  const response = await callGraphAPI(
    accessToken,
    'GET',
    `${config.getMailboxPrefix(mailbox)}/messages/${emailId}`,
    null,
    {
      $select: 'subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments',
      $expand: 'attachments($select=id,name,contentType,size)'
    }
  );

  let emailContent = `Subject: ${response.subject}\n`;
  emailContent += `From: ${response.from.emailAddress.name} <${response.from.emailAddress.address}>\n`;
  emailContent += `To: ${response.toRecipients.map(r => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(', ')}\n`;
  if (response.ccRecipients && response.ccRecipients.length > 0) {
    emailContent += `CC: ${response.ccRecipients.map(r => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(', ')}\n`;
  }
  emailContent += `Date: ${new Date(response.receivedDateTime).toLocaleString()}\n`;
  emailContent += `Has Attachments: ${response.hasAttachments ? 'Yes' : 'No'}\n\n`;
  emailContent += `Body:\n${response.body.content}`;

  // List attachment metadata (no download — use get_attachment to download specific files)
  if (response.attachments && response.attachments.length > 0) {
    emailContent += '\n\nAttachments:\n';

    for (let i = 0; i < response.attachments.length; i++) {
      const attachment = response.attachments[i];
      const sizeKB = attachment.size ? (attachment.size / 1024).toFixed(1) : 'Unknown';
      const odataType = attachment['@odata.type'] || '';

      if (odataType === '#microsoft.graph.referenceAttachment') {
        emailContent += `${i + 1}. ${attachment.name} (SharePoint/OneDrive link - ${sizeKB} KB)\n`;
        emailContent += `   Type: ${attachment.contentType || 'Unknown'}\n`;
        emailContent += `   Attachment ID: ${attachment.id}\n`;
      } else if (odataType === '#microsoft.graph.itemAttachment') {
        emailContent += `${i + 1}. ${attachment.name} (Outlook item)\n`;
        emailContent += `   Attachment ID: ${attachment.id}\n`;
      } else {
        // fileAttachment or unknown
        emailContent += `${i + 1}. ${attachment.name} (${sizeKB} KB)\n`;
        emailContent += `   Type: ${attachment.contentType || 'Unknown'}\n`;
        emailContent += `   Attachment ID: ${attachment.id}\n`;
      }
    }

    emailContent += '\nUse mail { operation: "get_attachment", emailId: "...", attachmentId: "..." } to download a specific attachment.';
  }

  return {
    content: [{ type: "text", text: emailContent }]
  };
}

async function getAttachment(accessToken, params) {
  const { emailId, attachmentId, mailbox } = params;

  // Clean up old attachments periodically
  try {
    cleanupOldAttachments();
  } catch (err) {
    console.error('Error during attachment cleanup:', err);
  }

  if (!emailId || !attachmentId) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameters: emailId and attachmentId. Read the email first to get attachment IDs."
      }]
    };
  }

  try {
    // Fetch single attachment with full content
    const attachment = await callGraphAPI(
      accessToken,
      'GET',
      `${config.getMailboxPrefix(mailbox)}/messages/${emailId}/attachments/${attachmentId}`,
      null,
      {}
    );

    const odataType = attachment['@odata.type'] || '';

    // Reference attachments (SharePoint/OneDrive links) — return URL for in-memory processing
    if (odataType === '#microsoft.graph.referenceAttachment') {
      const sourceUrl = attachment.sourceUrl || '';
      let result = `Attachment: ${attachment.name} (SharePoint/OneDrive link)\n`;
      result += `URL: ${sourceUrl}\n`;

      const localPath = convertSharePointUrlToLocal(sourceUrl);
      if (localPath) {
        result += `Local Path: ${localPath}\n`;
      }

      result += '\nThis is a cloud file link. Use files { operation: "search" } to find and download it via Graph API for in-memory processing.';
      return { content: [{ type: "text", text: result }] };
    }

    // File attachments (embedded) — download to temp
    if (odataType === '#microsoft.graph.fileAttachment') {
      const localPath = await downloadEmbeddedAttachment(attachment, emailId, accessToken);
      const sizeKB = attachment.size ? (attachment.size / 1024).toFixed(1) : 'Unknown';

      if (localPath) {
        return {
          content: [{
            type: "text",
            text: `Downloaded: ${attachment.name}\nType: ${attachment.contentType || 'Unknown'}\nSize: ${sizeKB} KB\nLocal Path: ${localPath}\n\nUse document tools (pandas, pandoc, pdftotext) to read this file.\nAfter processing, clean up with: mail { operation: "cleanup_attachment", path: "${localPath}" }`
          }]
        };
      }

      return {
        content: [{
          type: "text",
          text: `Failed to download attachment: ${attachment.name}`
        }]
      };
    }

    // Item attachments (Outlook items)
    if (odataType === '#microsoft.graph.itemAttachment') {
      let result = `Attachment: ${attachment.name} (Outlook item)\n`;
      if (attachment.item) {
        result += `Item Type: ${attachment.item['@odata.type']}\n`;
        if (attachment.item.subject) result += `Subject: ${attachment.item.subject}\n`;
        if (attachment.item.body) result += `\nContent:\n${attachment.item.body.content || '(empty)'}`;
      }
      return { content: [{ type: "text", text: result }] };
    }

    return {
      content: [{
        type: "text",
        text: `Unsupported attachment type: ${odataType}\nName: ${attachment.name}`
      }]
    };
  } catch (error) {
    console.error(`Error getting attachment:`, error);
    return {
      content: [{ type: "text", text: `Error getting attachment: ${error.message}` }]
    };
  }
}

function cleanupAttachment(params) {
  const { path: filePath } = params;

  if (!filePath) {
    return {
      content: [{ type: "text", text: "Missing required parameter: path" }]
    };
  }

  // Safety: only allow deleting files inside the temp attachments directory
  const tempDir = config.TEMP_ATTACHMENTS_PATH;
  const resolvedPath = path.resolve(filePath);
  const resolvedTempDir = path.resolve(tempDir);

  if (!resolvedPath.startsWith(resolvedTempDir + path.sep) && resolvedPath !== resolvedTempDir) {
    return {
      content: [{
        type: "text",
        text: `Security: can only clean up files in ${tempDir}. Provided path is outside the temp directory.`
      }]
    };
  }

  try {
    if (fs.existsSync(resolvedPath)) {
      fs.unlinkSync(resolvedPath);
      return {
        content: [{ type: "text", text: `Cleaned up: ${path.basename(resolvedPath)}` }]
      };
    }
    return {
      content: [{ type: "text", text: `File already removed: ${path.basename(resolvedPath)}` }]
    };
  } catch (err) {
    console.error('Error cleaning up attachment:', err);
    return {
      content: [{ type: "text", text: `Error cleaning up: ${err.message}` }]
    };
  }
}

async function sendEmail(accessToken, params) {
  try {
    if (!params || typeof params !== 'object') {
      return {
        content: [{
          type: "text",
          text: "Invalid parameters: expected an object with to, subject, and body"
        }]
      };
    }

    if (!params.to || !params.subject || !params.body) {
      return {
        content: [{
          type: "text",
          text: "Missing required parameters: to, subject, and body"
        }]
      };
    }

    const toRecipients = Array.isArray(params.to) ? params.to :
                         (typeof params.to === 'string' ? [params.to] : []);

    if (toRecipients.length === 0) {
      return {
        content: [{
          type: "text",
          text: "Invalid 'to' parameter. Please provide valid email address(es)."
        }]
      };
    }

    const message = {
      subject: params.subject,
      body: {
        contentType: "HTML",
        content: params.body
      },
      toRecipients: toRecipients.map(email => ({
        emailAddress: { address: email }
      }))
    };

    if (params.mailbox && params.mailbox !== 'me') {
      message.from = {
        emailAddress: { address: params.mailbox }
      };
    }

    if (params.cc) {
      const ccRecipients = Array.isArray(params.cc) ? params.cc : [params.cc];
      if (ccRecipients.length > 0) {
        message.ccRecipients = ccRecipients.map(email => ({
          emailAddress: { address: email }
        }));
      }
    }

    if (params.bcc) {
      const bccRecipients = Array.isArray(params.bcc) ? params.bcc : [params.bcc];
      if (bccRecipients.length > 0) {
        message.bccRecipients = bccRecipients.map(email => ({
          emailAddress: { address: email }
        }));
      }
    }

    const sendEndpoint = `${config.getMailboxPrefix(params.mailbox)}/sendMail`;
    await callGraphAPI(
      accessToken,
      'POST',
      sendEndpoint,
      {
        message: message,
        saveToSentItems: true
      },
      null
    );

    return {
      content: [{ type: "text", text: "Email sent successfully!" }]
    };
  } catch (error) {
    console.error('[EMAIL] Send error:', error.message);
    return {
      content: [{ type: "text", text: `Email send error: ${error.message}` }]
    };
  }
}

async function replyToEmail(accessToken, params) {
  try {
    if (!params || typeof params !== 'object') {
      return {
        content: [{
          type: "text",
          text: "Invalid parameters: expected an object with emailId and body"
        }]
      };
    }

    if (!params.emailId || !params.body) {
      return {
        content: [{
          type: "text",
          text: "Missing required parameters: emailId and body. Use 'reply' to respond in-thread to an existing email."
        }]
      };
    }

    // Use comment (not message.body) so Graph preserves the quoted original message.
    // message.body replaces the entire reply body, wiping the original.
    const replyPayload = {
      comment: params.body
    };

    // Only add message object if we need to override from/to/cc
    const messageOverrides = {};

    if (params.mailbox && params.mailbox !== 'me') {
      messageOverrides.from = {
        emailAddress: { address: params.mailbox }
      };
    }

    if (params.to) {
      const toRecipients = Array.isArray(params.to) ? params.to : [params.to];
      messageOverrides.toRecipients = toRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (params.cc) {
      const ccRecipients = Array.isArray(params.cc) ? params.cc : [params.cc];
      messageOverrides.ccRecipients = ccRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (Object.keys(messageOverrides).length > 0) {
      replyPayload.message = messageOverrides;
    }

    const replyEndpoint = `${config.getMailboxPrefix(params.mailbox)}/messages/${params.emailId}/reply`;
    await callGraphAPI(
      accessToken,
      'POST',
      replyEndpoint,
      replyPayload,
      null
    );

    return {
      content: [{ type: "text", text: "Reply sent successfully! The reply is threaded under the original email." }]
    };
  } catch (error) {
    console.error('[EMAIL] Reply error:', error.message);
    return {
      content: [{ type: "text", text: `Email reply error: ${error.message}` }]
    };
  }
}

// ============== DRAFT EMAIL FUNCTIONS ==============

function wrapHtmlBody(bodyContent) {
  // Strip CDATA wrappers if present
  if (bodyContent.includes('<![CDATA[')) {
    bodyContent = bodyContent.replace(/<!\[CDATA\[/g, '').replace(/\]\]>/g, '');
  }

  // Ensure proper HTML structure with professional font styling
  if (!bodyContent.includes('<html>')) {
    bodyContent = `<html>
<head>
<style>
body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, Arial, sans-serif;
  font-size: 11pt;
  color: #333333;
}
table {
  border-collapse: collapse;
  margin: 10px 0;
}
th, td {
  border: 1px solid #ddd;
  padding: 8px;
  text-align: left;
}
th {
  background-color: #f2f2f2;
  font-weight: bold;
}
h3 {
  color: #2c3e50;
  margin-top: 15px;
  margin-bottom: 10px;
}
ul, ol {
  margin: 10px 0;
}
</style>
</head>
<body>${bodyContent}</body>
</html>`;
  } else if (!bodyContent.includes('font-family') && !bodyContent.includes('Gulim')) {
    bodyContent = bodyContent.replace('<body>', `<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, Arial, sans-serif; font-size: 11pt; color: #333333;">`);
  }

  // Replace any Gulim font references with professional fonts
  bodyContent = bodyContent.replace(/font-family:\s*["']?Gulim["']?[^;]*/gi, "font-family: 'Segoe UI', Tahoma, Geneva, Verdana, Arial, sans-serif");

  return bodyContent;
}

async function createDraft(accessToken, params) {
  try {
    if (!params.subject && !params.body && !params.to) {
      return {
        content: [{
          type: "text",
          text: "At least one parameter required: subject, body, or to"
        }]
      };
    }

    const draftMessage = {};

    if (params.subject) {
      draftMessage.subject = params.subject;
    }

    if (params.body) {
      draftMessage.body = {
        contentType: "HTML",
        content: wrapHtmlBody(params.body)
      };
    }

    if (params.to) {
      const toRecipients = Array.isArray(params.to) ? params.to : [params.to];
      draftMessage.toRecipients = toRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (params.cc) {
      const ccRecipients = Array.isArray(params.cc) ? params.cc : [params.cc];
      draftMessage.ccRecipients = ccRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (params.bcc) {
      const bccRecipients = Array.isArray(params.bcc) ? params.bcc : [params.bcc];
      draftMessage.bccRecipients = bccRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    const response = await callGraphAPI(
      accessToken,
      'POST',
      `${config.getMailboxPrefix(params.mailbox)}/messages`,
      draftMessage,
      null
    );

    return {
      content: [{
        type: "text",
        text: `Draft created successfully!\nDraft ID: ${response.id}\nSubject: ${response.subject || '(No subject)'}`
      }]
    };
  } catch (error) {
    console.error('Error creating draft:', error);
    return {
      content: [{ type: "text", text: `Error creating draft: ${error.message}` }]
    };
  }
}

async function updateDraft(accessToken, params) {
  try {
    const { draftId, mailbox, ...updateParams } = params;

    if (!draftId) {
      return {
        content: [{
          type: "text",
          text: "Missing required parameter: draftId"
        }]
      };
    }

    const updateMessage = {};

    if (updateParams.subject) {
      updateMessage.subject = updateParams.subject;
    }

    if (updateParams.body) {
      updateMessage.body = {
        contentType: "HTML",
        content: wrapHtmlBody(updateParams.body)
      };
    }

    if (updateParams.to) {
      const toRecipients = Array.isArray(updateParams.to) ? updateParams.to : [updateParams.to];
      updateMessage.toRecipients = toRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (updateParams.cc) {
      const ccRecipients = Array.isArray(updateParams.cc) ? updateParams.cc : [updateParams.cc];
      updateMessage.ccRecipients = ccRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    if (updateParams.bcc) {
      const bccRecipients = Array.isArray(updateParams.bcc) ? updateParams.bcc : [updateParams.bcc];
      updateMessage.bccRecipients = bccRecipients.map(email => ({
        emailAddress: { address: email }
      }));
    }

    const response = await callGraphAPI(
      accessToken,
      'PATCH',
      `${config.getMailboxPrefix(mailbox)}/messages/${draftId}`,
      updateMessage,
      null
    );

    return {
      content: [{
        type: "text",
        text: `Draft updated successfully!\nDraft ID: ${response.id}\nSubject: ${response.subject || '(No subject)'}`
      }]
    };
  } catch (error) {
    console.error('Error updating draft:', error);
    return {
      content: [{ type: "text", text: `Error updating draft: ${error.message}` }]
    };
  }
}

async function sendDraft(accessToken, params) {
  try {
    const { draftId, mailbox } = params;

    if (!draftId) {
      return {
        content: [{
          type: "text",
          text: "Missing required parameter: draftId"
        }]
      };
    }

    await callGraphAPI(
      accessToken,
      'POST',
      `${config.getMailboxPrefix(mailbox)}/messages/${draftId}/send`,
      null,
      null
    );

    return {
      content: [{
        type: "text",
        text: "Draft sent successfully! The message has been moved to your Sent Items folder."
      }]
    };
  } catch (error) {
    console.error('Error sending draft:', error);
    return {
      content: [{ type: "text", text: `Error sending draft: ${error.message}` }]
    };
  }
}

async function listDrafts(accessToken, params) {
  try {
    const { maxResults = 10, mailbox } = params;

    const queryParams = {
      $top: maxResults,
      $select: config.EMAIL_SELECT_FIELDS,
      $orderby: 'lastModifiedDateTime desc'
    };

    const response = await callGraphAPI(
      accessToken,
      'GET',
      `${config.getMailboxPrefix(mailbox)}/mailFolders/drafts/messages`,
      null,
      queryParams
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No drafts found." }]
      };
    }

    const draftsList = response.value.map((draft, index) => {
      const toRecipients = draft.toRecipients?.map(r => r.emailAddress.address).join(', ') || '(No recipients)';
      const attachments = draft.hasAttachments ? ' 📎' : '';
      return `${index + 1}. ${draft.subject || '(No subject)'}${attachments}
   To: ${toRecipients}
   Modified: ${new Date(draft.lastModifiedDateTime).toLocaleString()}
   ID: ${draft.id}`;
    }).join('\n\n');

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} drafts:\n\n${draftsList}`
      }]
    };
  } catch (error) {
    console.error('Error listing drafts:', error);
    return {
      content: [{ type: "text", text: `Error listing drafts: ${error.message}` }]
    };
  }
}

// ============== UNIFIED MAIL ROUTER ==============

/**
 * Unified mail handler - single entry point for ALL email operations
 */
async function handleMail(args) {
  if (!args || typeof args !== 'object') {
    return {
      content: [{
        type: "text",
        text: "Invalid args: expected an object with 'operation' parameter"
      }]
    };
  }

  const { operation, ...params } = args;

  if (!operation) {
    return {
      content: [{
        type: "text",
        text: "Missing required parameter: operation. Valid operations: list, read, get_attachment, cleanup_attachment, send, reply, draft, update_draft, send_draft, list_drafts, search, move, list_folders, create_folder, list_rules, create_rule, focused, list_categories, create_category, apply_category, remove_category"
      }]
    };
  }

  try {
    switch (operation) {
      // Core email operations
      case 'list':
      case 'read':
      case 'get_attachment':
      case 'cleanup_attachment':
      case 'send':
      case 'reply':
      case 'draft':
      case 'update_draft':
      case 'send_draft':
      case 'list_drafts':
        return await handleEmail(args);

      // Search
      case 'search':
        return await handleEmailSearch(params);

      // Move
      case 'move':
        return await handleEmailMove(params);

      // Folder operations
      case 'list_folders':
        return await handleEmailFolder({ operation: 'list', ...params });
      case 'create_folder':
        return await handleEmailFolder({ operation: 'create', ...params });

      // Rules operations
      case 'list_rules':
        return await handleEmailRules({ operation: 'list', ...params });
      case 'create_rule':
        return await handleEmailRules({ operation: 'create', ...params });

      // Focused inbox
      case 'focused': {
        const accessToken = await ensureAuthenticated();
        return await getFocusedInbox(accessToken, params);
      }

      // Categories operations
      case 'list_categories':
        return await handleEmailCategories({ operation: 'list', ...params });
      case 'create_category':
        return await handleEmailCategories({ operation: 'create', ...params });
      case 'apply_category':
        return await handleEmailCategories({ operation: 'apply', ...params });
      case 'remove_category':
        return await handleEmailCategories({ operation: 'remove', ...params });

      default:
        return {
          content: [{
            type: "text",
            text: `Invalid operation: ${operation}. Valid operations: list, read, get_attachment, cleanup_attachment, send, reply, draft, update_draft, send_draft, list_drafts, search, move, list_folders, create_folder, list_rules, create_rule, focused, list_categories, create_category, apply_category, remove_category`
          }]
        };
    }
  } catch (error) {
    console.error(`Error in mail ${operation}:`, error);
    return {
      content: [{ type: "text", text: `Error in mail ${operation}: ${error.message}` }]
    };
  }
}

// Export consolidated single tool
const emailTools = [
  {
    name: "mail",
    description: "Unified email management: list, read, send, reply, draft, search, move, folders, rules, categories, and focused inbox",
    inputSchema: {
      type: "object",
      properties: {
        operation: {
          type: "string",
          enum: [
            "list", "read", "get_attachment", "cleanup_attachment",
            "send", "reply", "draft", "update_draft", "send_draft", "list_drafts",
            "search", "move",
            "list_folders", "create_folder",
            "list_rules", "create_rule",
            "focused",
            "list_categories", "create_category", "apply_category", "remove_category"
          ],
          description: "The operation to perform"
        },
        // Core email parameters
        emailId: { type: "string", description: "Email ID (for read, get_attachment, reply, move, apply_category)" },
        attachmentId: { type: "string", description: "Attachment ID (for get_attachment — get IDs from read)" },
        path: { type: "string", description: "Local file path (for cleanup_attachment)" },
        to: {
          type: "array",
          items: { type: "string" },
          description: "Recipient email addresses (for send/draft/reply)"
        },
        subject: { type: "string", description: "Email subject (for send/draft)" },
        body: { type: "string", description: "Email body in HTML format (for send/draft/reply)" },
        cc: {
          type: "array",
          items: { type: "string" },
          description: "CC recipients"
        },
        bcc: {
          type: "array",
          items: { type: "string" },
          description: "BCC recipients"
        },
        draftId: { type: "string", description: "Draft ID (for update_draft, send_draft)" },
        // Search parameters
        query: { type: "string", description: "Search text or KQL syntax (for search)" },
        from: { type: "string", description: "Filter by sender email (for search)" },
        hasAttachments: { type: "boolean", description: "Filter by attachments (for search)" },
        isRead: { type: "boolean", description: "Filter by read/unread (for search)" },
        importance: { type: "string", enum: ["high", "normal", "low"], description: "Filter by importance" },
        startDate: { type: "string", description: "Start date - ISO or relative 7d/1w/1m/1y (for search)" },
        endDate: { type: "string", description: "End date - ISO or relative (for search)" },
        folderName: { type: "string", description: "Folder name: inbox/sent/drafts/deleted/junk/archive (for search)" },
        useRelevance: { type: "boolean", description: "Sort by relevance (for search)" },
        includeDeleted: { type: "boolean", description: "Include deleted items (for search)" },
        // Move parameters
        emailIds: { type: "array", items: { type: "string" }, description: "Email IDs to move (for move)" },
        destinationFolderId: { type: "string", description: "Destination folder ID (for move)" },
        batch: { type: "boolean", description: "Use batch processing (for move)" },
        // Folder parameters
        folderId: { type: "string", description: "Folder ID (for list, search)" },
        displayName: { type: "string", description: "Name (for create_folder, create_rule, create_category)" },
        parentFolderId: { type: "string", description: "Parent folder ID (for create_folder)" },
        // Rule parameters
        enhanced: { type: "boolean", description: "Enhanced mode for rules" },
        fromAddresses: { type: "array", items: { type: "string" }, description: "From filter (for create_rule)" },
        moveToFolder: { type: "string", description: "Move-to folder ID (for create_rule)" },
        forwardTo: { type: "array", items: { type: "string" }, description: "Forward-to emails (for create_rule)" },
        subjectContains: { type: "array", items: { type: "string" }, description: "Subject keywords (for create_rule)" },
        // Category parameters
        categories: { type: "array", items: { type: "string" }, description: "Categories (for apply/remove)" },
        color: { type: "string", description: "Category color preset0-preset24 (for create_category)" },
        categoryId: { type: "string", description: "Category ID (for update/delete)" },
        // Common
        maxResults: { type: "number", description: "Maximum number of results" },
        mailbox: { type: "string", description: "Mailbox to access (default: authenticated user). Use for shared mailboxes." }
      },
      required: ["operation"]
    },
    handler: safeTool('mail', handleMail)
  }
];

module.exports = { emailTools };
