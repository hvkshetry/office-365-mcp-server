/**
 * HTML processing utilities for email body parsing
 */

/**
 * Strip HTML tags from text while preserving structure
 * @param {string} html - HTML content
 * @returns {string} - Plain text
 */
function stripHtml(html) {
  if (!html) return '';

  return html
    // Replace <br>, <p>, <div> with newlines
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    // Remove all other HTML tags
    .replace(/<[^>]+>/g, '')
    // Decode common HTML entities
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    // Clean up extra whitespace
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n\s*\n/g, '\n\n')
    .trim();
}

/**
 * Extract signature block from email body (English and French)
 * @param {string} bodyText - Plain text email body
 * @returns {string|null} - Signature text or null
 */
function extractSignature(bodyText) {
  if (!bodyText) return null;

  // Common signature indicators (English and French)
  const signatureMarkers = [
    // English
    'Best regards',
    'Regards',
    'Sincerely',
    'Thanks',
    'Cheers',
    'Best',
    'Thank you',
    // French
    'Cordialement',
    'Bien cordialement',
    'Salutations',
    'Salutations distinguées',
    'Respectueusement',
    'Amitiés',
    'Amicalement',
    'Bonne journée',
    'Bonne soirée',
    // Common
    '--',
    '___',
    'Sent from',
    'Envoyé depuis'
  ];

  // Find the last occurrence of any signature marker
  let lastMarkerIndex = -1;
  let markerFound = null;

  signatureMarkers.forEach(marker => {
    const index = bodyText.lastIndexOf(marker);
    if (index > lastMarkerIndex) {
      lastMarkerIndex = index;
      markerFound = marker;
    }
  });

  if (lastMarkerIndex === -1) {
    // No signature marker found, try to get last few lines
    const lines = bodyText.split('\n');
    if (lines.length > 3) {
      return lines.slice(-4).join('\n');
    }
    return null;
  }

  // Extract text from marker to end
  const signature = bodyText.substring(lastMarkerIndex);

  // Limit signature length (typically not more than 500 chars)
  return signature.length > 500 ? signature.substring(0, 500) : signature;
}

/**
 * Process HTML email body and extract plain text + signature
 * @param {string} htmlBody - HTML email body
 * @returns {Object} - {plainText, signature}
 */
function processHtmlBody(htmlBody) {
  if (!htmlBody) {
    return { plainText: '', signature: null };
  }

  const plainText = stripHtml(htmlBody);
  const signature = extractSignature(plainText);

  return {
    plainText,
    signature
  };
}

/**
 * Clean and normalize text for better parsing
 * @param {string} text - Text to clean
 * @returns {string} - Cleaned text
 */
function cleanText(text) {
  if (!text) return '';

  return text
    // Remove multiple spaces
    .replace(/\s+/g, ' ')
    // Remove leading/trailing whitespace
    .trim()
    // Normalize quotes
    .replace(/[""]/g, '"')
    .replace(/['']/g, "'");
}

/**
 * Check if email body likely contains a signature (English and French)
 * @param {string} bodyText - Plain text body
 * @returns {boolean} - True if signature markers found
 */
function hasSignature(bodyText) {
  if (!bodyText) return false;

  const signatureMarkers = [
    // English
    /best regards?/i,
    /sincerely/i,
    /thanks?/i,
    /cheers/i,
    /sent from/i,
    // French
    /cordialement/i,
    /bien cordialement/i,
    /salutations/i,
    /respectueusement/i,
    /amitiés/i,
    /amicalement/i,
    /bonne journée/i,
    /bonne soirée/i,
    /envoyé depuis/i,
    // Common
    /\n--\n/,
    /\n___+\n/
  ];

  return signatureMarkers.some(pattern => pattern.test(bodyText));
}

module.exports = {
  stripHtml,
  extractSignature,
  processHtmlBody,
  cleanText,
  hasSignature
};
