/**
 * Email Contact Extractor Tool
 * Extracts contact information from emails and exports to CSV
 */

const { ensureAuthenticated } = require('../auth');
const { callGraphAPI } = require('../utils/graph-api');
const { processHtmlBody, extractSignature } = require('../utils/html-processor');
const {
  extractContactFromEmailAddress,
  extractContactsFromBody,
  extractNameFromSignature
} = require('../utils/contact-parser');
const {
  deduplicateContacts,
  normalizeContact
} = require('../utils/deduplicator');
const {
  writeContactsToCSV,
  getDefaultCSVPath
} = require('../utils/csv-generator');
const config = require('../config');

/**
 * Parse date filter (supports ISO format or relative like "30d", "1w", "1m", "1y")
 * @param {string} dateStr - Date string
 * @returns {Date} - Parsed date
 */
function parseDateFilter(dateStr) {
  if (!dateStr) return null;

  // Try ISO format first
  const isoDate = new Date(dateStr);
  if (!isNaN(isoDate.getTime())) {
    return isoDate;
  }

  // Parse relative format
  const match = dateStr.match(/^(\d+)([dwmy])$/);
  if (match) {
    const amount = parseInt(match[1]);
    const unit = match[2];
    const now = new Date();

    switch (unit) {
      case 'd':
        now.setDate(now.getDate() - amount);
        break;
      case 'w':
        now.setDate(now.getDate() - (amount * 7));
        break;
      case 'm':
        now.setMonth(now.getMonth() - amount);
        break;
      case 'y':
        now.setFullYear(now.getFullYear() - amount);
        break;
    }

    return now;
  }

  return null;
}

/**
 * Fetch emails from inbox with filters
 * @param {string} accessToken - Access token
 * @param {Object} params - Filter parameters
 * @returns {Promise<Array>} - Array of emails
 */
async function fetchEmails(accessToken, params) {
  const {
    searchQuery = null,
    maxEmails = 1000,
    startDate = null,
    endDate = null
  } = params;

  const emails = [];
  let currentUrl = 'me/messages';
  const queryParams = {
    $top: Math.min(maxEmails, 100), // Batch size
    $select: 'id,from,toRecipients,ccRecipients,subject,bodyPreview,body,receivedDateTime'
  };

  // Add search query if provided
  if (searchQuery) {
    queryParams.$search = `"${searchQuery}"`;
  }

  // Add date filters
  const filters = [];
  if (startDate) {
    const date = parseDateFilter(startDate);
    if (date) {
      filters.push(`receivedDateTime ge ${date.toISOString()}`);
    }
  }
  if (endDate) {
    const date = parseDateFilter(endDate);
    if (date) {
      filters.push(`receivedDateTime le ${date.toISOString()}`);
    }
  }

  if (filters.length > 0) {
    queryParams.$filter = filters.join(' and ');
  }

  console.error(`Fetching emails with query params:`, queryParams);

  // Fetch emails with pagination
  while (emails.length < maxEmails) {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      currentUrl,
      null,
      queryParams
    );

    if (!response.value || response.value.length === 0) {
      break;
    }

    emails.push(...response.value);

    // Check if we have more pages
    if (response['@odata.nextLink'] && emails.length < maxEmails) {
      // Extract the next page URL
      const nextUrl = new URL(response['@odata.nextLink']);
      currentUrl = nextUrl.pathname.replace('/v1.0/', '') + nextUrl.search;
      // Clear query params since they're in the URL
      Object.keys(queryParams).forEach(key => delete queryParams[key]);
    } else {
      break;
    }
  }

  return emails.slice(0, maxEmails);
}

/**
 * Extract contacts from a single email
 * @param {Object} email - Email object from Graph API
 * @param {boolean} includeBody - Whether to parse body
 * @returns {Array<Object>} - Array of contacts found in email
 */
function extractContactsFromEmail(email, includeBody = true) {
  const contacts = [];
  const seenEmails = new Set();

  // Extract from metadata (from, to, cc)
  const metadataContacts = [];

  if (email.from && email.from.emailAddress) {
    metadataContacts.push(email.from.emailAddress);
  }

  if (Array.isArray(email.toRecipients)) {
    email.toRecipients.forEach(recipient => {
      if (recipient.emailAddress) {
        metadataContacts.push(recipient.emailAddress);
      }
    });
  }

  if (Array.isArray(email.ccRecipients)) {
    email.ccRecipients.forEach(recipient => {
      if (recipient.emailAddress) {
        metadataContacts.push(recipient.emailAddress);
      }
    });
  }

  // Process metadata contacts
  metadataContacts.forEach(emailAddr => {
    const contact = extractContactFromEmailAddress(emailAddr);
    if (contact && contact.email && !seenEmails.has(contact.email)) {
      contact.firstSeenDate = email.receivedDateTime || new Date().toISOString();
      contact.extractionConfidence = 'high'; // Metadata is high confidence
      contacts.push(contact);
      seenEmails.add(contact.email);
    }
  });

  // Extract from body if requested
  if (includeBody && email.body && email.body.content) {
    const { plainText, signature } = processHtmlBody(email.body.content);

    // Extract contacts from signature
    if (signature) {
      const name = extractNameFromSignature(signature);
      const bodyContacts = extractContactsFromBody(signature);

      if (bodyContacts.emails && bodyContacts.emails.length > 0) {
        bodyContacts.emails.forEach(emailAddr => {
          if (!seenEmails.has(emailAddr)) {
            const contact = {
              email: emailAddr,
              displayName: name,
              phoneNumbers: bodyContacts.phoneNumbers || [],
              linkedInUrls: bodyContacts.linkedInUrls || [],
              companyName: bodyContacts.companyName,
              jobTitle: bodyContacts.jobTitle,
              source: 'signature',
              firstSeenDate: email.receivedDateTime || new Date().toISOString(),
              extractionConfidence: 'medium'
            };

            // Parse display name if available
            if (name) {
              const nameParts = name.split(/\s+/);
              if (nameParts.length >= 2) {
                contact.firstName = nameParts[0];
                contact.lastName = nameParts.slice(1).join(' ');
              }
            }

            contacts.push(contact);
            seenEmails.add(emailAddr);
          }
        });
      }
    }

    // Extract from full body text (lower priority)
    const fullBodyContacts = extractContactsFromBody(plainText);
    if (fullBodyContacts.emails && fullBodyContacts.emails.length > 0) {
      fullBodyContacts.emails.forEach(emailAddr => {
        if (!seenEmails.has(emailAddr)) {
          const contact = {
            email: emailAddr,
            phoneNumbers: fullBodyContacts.phoneNumbers || [],
            linkedInUrls: fullBodyContacts.linkedInUrls || [],
            companyName: fullBodyContacts.companyName,
            jobTitle: fullBodyContacts.jobTitle,
            source: 'body',
            firstSeenDate: email.receivedDateTime || new Date().toISOString(),
            extractionConfidence: 'low'
          };

          contacts.push(contact);
          seenEmails.add(emailAddr);
        }
      });
    }
  }

  return contacts;
}

/**
 * Fetch all Outlook contacts for cross-referencing
 * @param {string} accessToken - Access token
 * @returns {Promise<Set>} - Set of email addresses in Outlook
 */
async function fetchOutlookContactEmails(accessToken) {
  const outlookEmails = new Set();
  let currentUrl = 'me/contacts';
  const queryParams = {
    $select: 'emailAddresses',
    $top: 100
  };

  console.error('Fetching Outlook contacts for cross-reference...');

  while (true) {
    try {
      const response = await callGraphAPI(
        accessToken,
        'GET',
        currentUrl,
        null,
        queryParams
      );

      if (!response.value || response.value.length === 0) {
        break;
      }

      response.value.forEach(contact => {
        if (Array.isArray(contact.emailAddresses)) {
          contact.emailAddresses.forEach(emailObj => {
            if (emailObj.address) {
              outlookEmails.add(emailObj.address.toLowerCase().trim());
            }
          });
        }
      });

      // Check for next page
      if (response['@odata.nextLink']) {
        const nextUrl = new URL(response['@odata.nextLink']);
        currentUrl = nextUrl.pathname.replace('/v1.0/', '') + nextUrl.search;
        Object.keys(queryParams).forEach(key => delete queryParams[key]);
      } else {
        break;
      }
    } catch (error) {
      console.error('Error fetching Outlook contacts:', error.message);
      break;
    }
  }

  console.error(`Found ${outlookEmails.size} email addresses in Outlook contacts`);
  return outlookEmails;
}

/**
 * Main handler for contact extraction
 * @param {Object} args - Tool arguments
 * @returns {Promise<Object>} - Tool response
 */
async function extractContactsFromEmails(args) {
  const {
    searchQuery = null,
    maxEmails = 1000,
    includeBody = true,
    outputPath = null,
    startDate = null,
    endDate = null
  } = args;

  try {
    const accessToken = await ensureAuthenticated();

    console.error('Starting contact extraction...');
    console.error(`Parameters: maxEmails=${maxEmails}, includeBody=${includeBody}`);

    // Fetch emails
    const emails = await fetchEmails(accessToken, {
      searchQuery,
      maxEmails,
      startDate,
      endDate
    });

    console.error(`Fetched ${emails.length} emails`);

    if (emails.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No emails found matching the criteria."
        }]
      };
    }

    // Extract contacts from all emails
    let allContacts = [];
    emails.forEach((email, index) => {
      if (index % 100 === 0) {
        console.error(`Processing email ${index + 1}/${emails.length}`);
      }

      const emailContacts = extractContactsFromEmail(email, includeBody);
      allContacts.push(...emailContacts);
    });

    console.error(`Extracted ${allContacts.length} contacts (before deduplication)`);

    // Normalize contacts
    allContacts = allContacts
      .map(contact => normalizeContact(contact))
      .filter(contact => contact !== null);

    // Deduplicate
    const { deduplicated, duplicatesRemoved } = deduplicateContacts(allContacts);

    console.error(`After deduplication: ${deduplicated.length} unique contacts`);

    // Fetch Outlook contacts for cross-reference
    const outlookEmails = await fetchOutlookContactEmails(accessToken);

    // Mark contacts as new or existing
    deduplicated.forEach(contact => {
      contact.isInOutlook = outlookEmails.has(contact.email);
    });

    const newContacts = deduplicated.filter(c => !c.isInOutlook);

    console.error(`Found ${newContacts.length} new contacts not in Outlook`);

    // Calculate stats
    const stats = {
      emailsWithLinkedIn: deduplicated.filter(c => c.linkedInUrls && c.linkedInUrls.length > 0).length,
      emailsWithPhones: deduplicated.filter(c => c.phoneNumbers && c.phoneNumbers.length > 0).length,
      totalDuplicatesRemoved: duplicatesRemoved
    };

    // Write to CSV
    const csvPath = outputPath || getDefaultCSVPath();
    const absolutePath = await writeContactsToCSV(deduplicated, csvPath);

    console.error(`CSV written to: ${absolutePath}`);

    // Build result message
    const resultText = `Contact extraction completed successfully!

📊 **Summary:**
- Total emails processed: ${emails.length}
- Unique contacts found: ${deduplicated.length}
- New contacts (not in Outlook): ${newContacts.length}
- Contacts with LinkedIn: ${stats.emailsWithLinkedIn}
- Contacts with phone numbers: ${stats.emailsWithPhones}
- Duplicates removed: ${stats.totalDuplicatesRemoved}

📄 **CSV exported to:** ${absolutePath}

The CSV file contains all extracted contacts with the following fields:
- email, displayName, firstName, lastName
- phoneNumbers, linkedInUrls
- companyName, jobTitle
- source, isInOutlook, firstSeenDate, extractionConfidence`;

    return {
      content: [{
        type: "text",
        text: resultText
      }]
    };

  } catch (error) {
    console.error('Error in extractContactsFromEmails:', error);
    return {
      content: [{
        type: "text",
        text: `Error extracting contacts: ${error.message}`
      }],
      isError: true
    };
  }
}

module.exports = {
  extractContactsFromEmails,
  fetchEmails,
  extractContactsFromEmail,
  fetchOutlookContactEmails
};
