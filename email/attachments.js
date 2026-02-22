/**
 * Email attachment helpers - download, SharePoint URL mapping, cleanup
 */

const { callGraphAPI } = require('../utils/graph-api');
const config = require('../config');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

/**
 * Helper function to convert SharePoint URLs to local sync paths
 * @param {string} sharePointUrl - The SharePoint URL to convert
 * @returns {string|null} - The local path or null if unable to map
 */
function convertSharePointUrlToLocal(sharePointUrl) {
  try {
    // Parse the SharePoint URL
    const url = new URL(sharePointUrl);
    const pathParts = url.pathname.split('/');

    // Look for sites index
    const sitesIndex = pathParts.indexOf('sites');
    if (sitesIndex === -1) {
      // Check for personal OneDrive format
      if (url.hostname.includes('-my.sharepoint.com')) {
        // Personal OneDrive format: https://circleh2o-my.sharepoint.com/personal/user_domain/...
        const personalIndex = pathParts.indexOf('personal');
        if (personalIndex !== -1 && pathParts.length > personalIndex + 2) {
          const remainingPath = pathParts.slice(personalIndex + 2).join('/');
          return `${config.ONEDRIVE_SYNC_PATH}/${decodeURIComponent(remainingPath)}`;
        }
      }
      return null;
    }

    // Get site name (e.g., "Proposals")
    const siteName = pathParts[sitesIndex + 1];

    // Find "Shared Documents" or "Documents" index
    let docsIndex = pathParts.indexOf('Shared Documents');
    if (docsIndex === -1) {
      docsIndex = pathParts.indexOf('Documents');
    }

    if (docsIndex === -1) {
      return null;
    }

    // Get remaining path after documents folder
    const docPath = pathParts.slice(docsIndex + 1).join('/');

    // Construct local path
    // Pattern: [SHAREPOINT_SYNC_PATH]/[Site] - Documents/[remaining path]
    const localPath = `${config.SHAREPOINT_SYNC_PATH}/${siteName} - Documents/${decodeURIComponent(docPath)}`;

    return localPath;
  } catch (err) {
    console.error('Error parsing SharePoint URL:', err);
    return null;
  }
}

/**
 * Download and save embedded attachment to local temp directory
 * @param {object} attachment - The attachment object from Graph API
 * @param {string} emailId - The email ID for API calls
 * @param {string} accessToken - Auth token for Graph API
 * @returns {string|null} - Local file path or null on error
 */
async function downloadEmbeddedAttachment(attachment, emailId, accessToken) {
  try {
    const tempDir = config.TEMP_ATTACHMENTS_PATH;

    // Ensure directory exists
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    // Generate unique filename to avoid conflicts
    const timestamp = Date.now();
    const hash = crypto.createHash('md5').update(attachment.id).digest('hex').substring(0, 8);
    const sanitizedName = attachment.name.replace(/[^a-zA-Z0-9.-]/g, '_');
    const fileName = `${timestamp}_${hash}_${sanitizedName}`;
    const filePath = path.join(tempDir, fileName);

    let fileData;

    // Check if contentBytes is available (small files < 3MB)
    if (attachment.contentBytes) {
      // Use existing base64 data
      fileData = Buffer.from(attachment.contentBytes, 'base64');
    } else {
      // Large file - fetch using /$value endpoint
      const binaryData = await callGraphAPI(
        accessToken,
        'GET',
        `me/messages/${emailId}/attachments/${attachment.id}/$value`,
        null,
        {}
      );

      // callGraphAPI returns raw data for non-JSON responses
      fileData = Buffer.from(binaryData, 'binary');
    }

    // Save to file
    fs.writeFileSync(filePath, fileData);

    console.error(`Attachment saved to: ${filePath}`);
    return filePath;

  } catch (err) {
    console.error('Error downloading attachment:', err);
    return null;
  }
}

/**
 * Clean up old attachment files (older than 24 hours)
 */
function cleanupOldAttachments() {
  const tempDir = config.TEMP_ATTACHMENTS_PATH;
  const maxAge = 1 * 60 * 60 * 1000; // 1 hour in milliseconds

  if (!fs.existsSync(tempDir)) return;

  const files = fs.readdirSync(tempDir);
  const now = Date.now();

  files.forEach(file => {
    const filePath = path.join(tempDir, file);
    const stats = fs.statSync(filePath);
    const age = now - stats.mtimeMs;

    if (age > maxAge) {
      try {
        fs.unlinkSync(filePath);
        console.error(`Cleaned up old attachment: ${file}`);
      } catch (err) {
        console.error(`Failed to delete old attachment: ${file}`, err);
      }
    }
  });
}

module.exports = {
  convertSharePointUrlToLocal,
  downloadEmbeddedAttachment,
  cleanupOldAttachments
};
