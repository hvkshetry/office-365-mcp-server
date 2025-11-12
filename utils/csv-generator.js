/**
 * CSV generation utilities for contact export
 */

const fs = require('fs');
const path = require('path');

/**
 * Escape CSV field value
 * @param {string} value - Field value
 * @returns {string} - Escaped value
 */
function escapeCSVField(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const stringValue = String(value);

  // If value contains comma, quote, or newline, wrap in quotes and escape quotes
  if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }

  return stringValue;
}

/**
 * Convert array to CSV string
 * @param {Array} arr - Array to convert
 * @param {string} separator - Separator (default: semicolon)
 * @returns {string} - CSV string
 */
function arrayToCSVString(arr, separator = ';') {
  if (!Array.isArray(arr) || arr.length === 0) {
    return '';
  }
  return arr.join(separator);
}

/**
 * Generate CSV content from contacts array
 * @param {Array<Object>} contacts - Array of contact objects
 * @returns {string} - CSV content
 */
function generateCSV(contacts) {
  if (!Array.isArray(contacts) || contacts.length === 0) {
    return '';
  }

  // CSV Headers
  const headers = [
    'email',
    'displayName',
    'firstName',
    'lastName',
    'phoneNumbers',
    'linkedInUrls',
    'companyName',
    'jobTitle',
    'source',
    'isInOutlook',
    'firstSeenDate',
    'extractionConfidence'
  ];

  // Build CSV rows
  const rows = [headers.join(',')];

  contacts.forEach(contact => {
    const row = [
      escapeCSVField(contact.email),
      escapeCSVField(contact.displayName),
      escapeCSVField(contact.firstName),
      escapeCSVField(contact.lastName),
      escapeCSVField(arrayToCSVString(contact.phoneNumbers)),
      escapeCSVField(arrayToCSVString(contact.linkedInUrls)),
      escapeCSVField(contact.companyName),
      escapeCSVField(contact.jobTitle),
      escapeCSVField(contact.source),
      escapeCSVField(contact.isInOutlook),
      escapeCSVField(contact.firstSeenDate),
      escapeCSVField(contact.extractionConfidence)
    ];

    rows.push(row.join(','));
  });

  return rows.join('\n');
}

/**
 * Write contacts to CSV file
 * @param {Array<Object>} contacts - Array of contact objects
 * @param {string} filePath - Output file path
 * @returns {Promise<string>} - Absolute path to created file
 */
async function writeContactsToCSV(contacts, filePath) {
  if (!filePath) {
    throw new Error('File path is required');
  }

  // Ensure directory exists
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  // Generate CSV content
  const csvContent = generateCSV(contacts);

  // Write to file
  fs.writeFileSync(filePath, csvContent, 'utf8');

  // Return absolute path
  return path.resolve(filePath);
}

/**
 * Get default CSV path
 * @returns {string} - Default CSV file path
 */
function getDefaultCSVPath() {
  const homeDir = process.env.HOME || process.env.USERPROFILE || '/tmp';
  return path.join(homeDir, 'extracted_contacts.csv');
}

module.exports = {
  escapeCSVField,
  arrayToCSVString,
  generateCSV,
  writeContactsToCSV,
  getDefaultCSVPath
};
