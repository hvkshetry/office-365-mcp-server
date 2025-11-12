/**
 * Contact deduplication utilities
 */

/**
 * Merge two contact objects, preferring more complete information
 * @param {Object} contact1 - First contact
 * @param {Object} contact2 - Second contact
 * @returns {Object} - Merged contact
 */
function mergeContacts(contact1, contact2) {
  // Prefer metadata over body-extracted data
  const sourcePreference = { metadata: 3, signature: 2, body: 1 };

  const getSourcePriority = (source) => sourcePreference[source] || 0;

  const merged = { ...contact1 };

  // Merge each field, preferring non-null values with higher source priority
  Object.keys(contact2).forEach(key => {
    if (key === 'email') return; // Email is the primary key, don't change

    const val1 = contact1[key];
    const val2 = contact2[key];

    // If contact2 has a value and contact1 doesn't, use contact2's value
    if (val2 && !val1) {
      merged[key] = val2;
      return;
    }

    // If both have values, prefer based on source priority
    if (val1 && val2) {
      const priority1 = getSourcePriority(contact1.source);
      const priority2 = getSourcePriority(contact2.source);

      if (priority2 > priority1) {
        merged[key] = val2;
      }
    }
  });

  // Merge arrays (phone numbers, etc.)
  if (Array.isArray(contact1.phoneNumbers) && Array.isArray(contact2.phoneNumbers)) {
    merged.phoneNumbers = [...new Set([...contact1.phoneNumbers, ...contact2.phoneNumbers])];
  }

  if (Array.isArray(contact1.linkedInUrls) && Array.isArray(contact2.linkedInUrls)) {
    merged.linkedInUrls = [...new Set([...contact1.linkedInUrls, ...contact2.linkedInUrls])];
  }

  // Use the higher confidence
  if (contact2.extractionConfidence && contact1.extractionConfidence) {
    const confidenceLevels = { high: 3, medium: 2, low: 1 };
    const conf1 = confidenceLevels[contact1.extractionConfidence] || 0;
    const conf2 = confidenceLevels[contact2.extractionConfidence] || 0;

    if (conf2 > conf1) {
      merged.extractionConfidence = contact2.extractionConfidence;
    }
  }

  // Keep earliest first seen date
  if (contact1.firstSeenDate && contact2.firstSeenDate) {
    const date1 = new Date(contact1.firstSeenDate);
    const date2 = new Date(contact2.firstSeenDate);
    merged.firstSeenDate = date1 < date2 ? contact1.firstSeenDate : contact2.firstSeenDate;
  }

  return merged;
}

/**
 * Deduplicate array of contacts by email address
 * @param {Array<Object>} contacts - Array of contact objects
 * @returns {Object} - {deduplicated: Array, duplicatesRemoved: number}
 */
function deduplicateContacts(contacts) {
  if (!Array.isArray(contacts) || contacts.length === 0) {
    return { deduplicated: [], duplicatesRemoved: 0 };
  }

  const contactMap = new Map();

  contacts.forEach(contact => {
    if (!contact.email) return;

    const email = contact.email.toLowerCase().trim();

    if (contactMap.has(email)) {
      // Merge with existing contact
      const existing = contactMap.get(email);
      const merged = mergeContacts(existing, contact);
      contactMap.set(email, merged);
    } else {
      // Add new contact
      contactMap.set(email, contact);
    }
  });

  const deduplicated = Array.from(contactMap.values());
  const duplicatesRemoved = contacts.length - deduplicated.length;

  return {
    deduplicated,
    duplicatesRemoved
  };
}

/**
 * Calculate extraction confidence based on available data
 * @param {Object} contact - Contact object
 * @returns {string} - 'high', 'medium', or 'low'
 */
function calculateConfidence(contact) {
  let score = 0;

  // Email is required (already have it)
  score += 1;

  // High value fields
  if (contact.displayName && contact.displayName.length > 0) score += 2;
  if (contact.firstName && contact.lastName) score += 2;

  // Medium value fields
  if (contact.phoneNumbers && contact.phoneNumbers.length > 0) score += 1;
  if (contact.linkedInUrls && contact.linkedInUrls.length > 0) score += 1;
  if (contact.companyName) score += 1;
  if (contact.jobTitle) score += 1;

  // Source quality
  if (contact.source === 'metadata') score += 1;
  else if (contact.source === 'signature') score += 0.5;

  // Determine confidence level
  if (score >= 7) return 'high';
  if (score >= 4) return 'medium';
  return 'low';
}

/**
 * Normalize contact data
 * @param {Object} contact - Raw contact object
 * @returns {Object} - Normalized contact
 */
function normalizeContact(contact) {
  if (!contact.email) return null;

  const normalized = {
    email: contact.email.toLowerCase().trim(),
    displayName: contact.displayName?.trim() || null,
    firstName: contact.firstName?.trim() || null,
    lastName: contact.lastName?.trim() || null,
    phoneNumbers: Array.isArray(contact.phoneNumbers) ? contact.phoneNumbers :
                   (contact.phoneNumber ? [contact.phoneNumber] : []),
    linkedInUrls: Array.isArray(contact.linkedInUrls) ? contact.linkedInUrls :
                   (contact.linkedInUrl ? [contact.linkedInUrl] : []),
    companyName: contact.companyName?.trim() || null,
    jobTitle: contact.jobTitle?.trim() || null,
    source: contact.source || 'unknown',
    isInOutlook: contact.isInOutlook || false,
    firstSeenDate: contact.firstSeenDate || new Date().toISOString(),
    extractionConfidence: contact.extractionConfidence || 'low'
  };

  // Calculate confidence if not provided
  if (!contact.extractionConfidence) {
    normalized.extractionConfidence = calculateConfidence(normalized);
  }

  return normalized;
}

module.exports = {
  mergeContacts,
  deduplicateContacts,
  calculateConfidence,
  normalizeContact
};
