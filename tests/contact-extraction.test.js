/**
 * Tests for email contact extraction feature
 */

const { describe, it, expect, beforeEach } = require('@jest/globals');
const {
  extractEmails,
  extractPhoneNumbers,
  extractLinkedInUrls,
  extractNameFromSignature,
  extractCompanyName,
  extractJobTitle,
  parseDisplayName,
  extractContactsFromBody,
  extractContactFromEmailAddress
} = require('../utils/contact-parser');

const {
  stripHtml,
  extractSignature,
  processHtmlBody,
  hasSignature
} = require('../utils/html-processor');

const {
  mergeContacts,
  deduplicateContacts,
  calculateConfidence,
  normalizeContact
} = require('../utils/deduplicator');

const {
  escapeCSVField,
  arrayToCSVString,
  generateCSV
} = require('../utils/csv-generator');

const {
  extractContactsFromEmail
} = require('../tools/email-contact-extractor');

describe('Contact Parser', () => {
  describe('extractEmails', () => {
    it('should extract email addresses from text', () => {
      const text = 'Contact me at john@example.com or jane.doe@company.org';
      const emails = extractEmails(text);
      expect(emails).toContain('john@example.com');
      expect(emails).toContain('jane.doe@company.org');
      expect(emails.length).toBe(2);
    });

    it('should filter out image files', () => {
      const text = 'Email: john@example.com and logo.png@server.com';
      const emails = extractEmails(text);
      expect(emails).toContain('john@example.com');
      expect(emails).not.toContain('logo.png@server.com');
    });

    it('should filter out noreply addresses', () => {
      const text = 'Contact: john@example.com or noreply@service.com';
      const emails = extractEmails(text);
      expect(emails).toContain('john@example.com');
      expect(emails).not.toContain('noreply@service.com');
    });
  });

  describe('extractPhoneNumbers', () => {
    it('should extract US phone numbers in various formats', () => {
      const text = 'Call me at (555) 123-4567 or 555-987-6543';
      const phones = extractPhoneNumbers(text);
      expect(phones.length).toBeGreaterThan(0);
    });

    it('should extract international phone numbers', () => {
      const text = 'Phone: +1-555-123-4567';
      const phones = extractPhoneNumbers(text);
      expect(phones.length).toBeGreaterThan(0);
    });

    it('should extract French phone numbers', () => {
      const text = 'Appelez-moi au +33 1 23 45 67 89 ou 06.12.34.56.78';
      const phones = extractPhoneNumbers(text);
      expect(phones.length).toBeGreaterThan(0);
    });

    it('should extract French mobile numbers', () => {
      const text = 'Mobile: 06 12 34 56 78';
      const phones = extractPhoneNumbers(text);
      expect(phones.length).toBeGreaterThan(0);
    });
  });

  describe('extractLinkedInUrls', () => {
    it('should extract LinkedIn profile URLs', () => {
      const text = 'Connect with me: https://www.linkedin.com/in/johndoe or linkedin.com/in/janedoe';
      const urls = extractLinkedInUrls(text);
      expect(urls.length).toBe(2);
      expect(urls[0]).toContain('linkedin.com/in/');
    });

    it('should normalize LinkedIn URLs', () => {
      const text = 'Profile: linkedin.com/in/johndoe';
      const urls = extractLinkedInUrls(text);
      expect(urls[0]).toMatch(/^https:\/\//);
    });
  });

  describe('parseDisplayName', () => {
    it('should parse first and last name', () => {
      const { firstName, lastName } = parseDisplayName('John Doe');
      expect(firstName).toBe('John');
      expect(lastName).toBe('Doe');
    });

    it('should handle multi-part last names', () => {
      const { firstName, lastName } = parseDisplayName('John Van Der Berg');
      expect(firstName).toBe('John');
      expect(lastName).toBe('Van Der Berg');
    });

    it('should handle single names', () => {
      const { firstName, lastName } = parseDisplayName('Madonna');
      expect(firstName).toBe('Madonna');
      expect(lastName).toBeNull();
    });
  });

  describe('extractContactFromEmailAddress', () => {
    it('should extract contact from email address object', () => {
      const emailAddr = {
        name: 'John Doe',
        address: 'john@example.com'
      };
      const contact = extractContactFromEmailAddress(emailAddr);
      expect(contact.email).toBe('john@example.com');
      expect(contact.displayName).toBe('John Doe');
      expect(contact.firstName).toBe('John');
      expect(contact.lastName).toBe('Doe');
      expect(contact.source).toBe('metadata');
    });
  });
});

describe('HTML Processor', () => {
  describe('stripHtml', () => {
    it('should remove HTML tags', () => {
      const html = '<p>Hello <b>World</b></p>';
      const text = stripHtml(html);
      expect(text).toBe('Hello World');
    });

    it('should preserve line breaks', () => {
      const html = '<p>Line 1</p><p>Line 2</p>';
      const text = stripHtml(html);
      expect(text).toContain('Line 1');
      expect(text).toContain('Line 2');
    });

    it('should decode HTML entities', () => {
      const html = 'AT&amp;T &lt;company&gt;';
      const text = stripHtml(html);
      expect(text).toBe('AT&T <company>');
    });
  });

  describe('extractSignature', () => {
    it('should extract signature from text', () => {
      const body = 'Email content here.\n\nBest regards,\nJohn Doe\nCEO, Example Corp';
      const signature = extractSignature(body);
      expect(signature).toContain('Best regards');
      expect(signature).toContain('John Doe');
    });

    it('should extract French signature from text', () => {
      const body = 'Contenu de l\'email ici.\n\nCordialement,\nJean Dupont\nDirecteur, Société Exemple';
      const signature = extractSignature(body);
      expect(signature).toContain('Cordialement');
      expect(signature).toContain('Jean Dupont');
    });

    it('should return null for short emails without signature marker', () => {
      const body = 'Just a simple email with no signature';
      const signature = extractSignature(body);
      expect(signature).toBeNull();
    });

    it('should return last few lines for longer emails without signature marker', () => {
      const body = 'Line 1\nLine 2\nLine 3\nLine 4\nLine 5';
      const signature = extractSignature(body);
      expect(signature).toBeTruthy();
      expect(signature).toContain('Line 5');
    });
  });

  describe('hasSignature', () => {
    it('should detect signature markers', () => {
      const body = 'Email content\n\nBest regards,\nJohn';
      expect(hasSignature(body)).toBe(true);
    });

    it('should detect French signature markers', () => {
      const body = 'Contenu de l\'email\n\nCordialement,\nJean';
      expect(hasSignature(body)).toBe(true);
    });

    it('should detect Bien cordialement', () => {
      const body = 'Message\n\nBien cordialement,\nMarie';
      expect(hasSignature(body)).toBe(true);
    });

    it('should return false if no signature', () => {
      const body = 'Just plain text';
      expect(hasSignature(body)).toBe(false);
    });
  });
});

describe('Deduplicator', () => {
  describe('mergeContacts', () => {
    it('should merge two contacts with same email', () => {
      const contact1 = {
        email: 'john@example.com',
        displayName: 'John',
        source: 'metadata'
      };
      const contact2 = {
        email: 'john@example.com',
        phoneNumbers: ['555-1234'],
        linkedInUrls: ['https://linkedin.com/in/john'],
        source: 'body'
      };
      const merged = mergeContacts(contact1, contact2);
      expect(merged.displayName).toBe('John'); // Prefer metadata
      expect(merged.phoneNumbers).toEqual(['555-1234']);
      expect(merged.linkedInUrls).toEqual(['https://linkedin.com/in/john']);
    });

    it('should prefer metadata source over body source', () => {
      const contact1 = {
        email: 'john@example.com',
        displayName: 'John Doe',
        source: 'metadata'
      };
      const contact2 = {
        email: 'john@example.com',
        displayName: 'J. Doe',
        source: 'body'
      };
      const merged = mergeContacts(contact1, contact2);
      expect(merged.displayName).toBe('John Doe');
    });
  });

  describe('deduplicateContacts', () => {
    it('should remove duplicate contacts', () => {
      const contacts = [
        { email: 'john@example.com', displayName: 'John' },
        { email: 'jane@example.com', displayName: 'Jane' },
        { email: 'john@example.com', phoneNumbers: ['555-1234'] }
      ];
      const { deduplicated, duplicatesRemoved } = deduplicateContacts(contacts);
      expect(deduplicated.length).toBe(2);
      expect(duplicatesRemoved).toBe(1);
    });

    it('should merge data from duplicates', () => {
      const contacts = [
        { email: 'john@example.com', displayName: 'John', source: 'metadata' },
        { email: 'john@example.com', phoneNumbers: ['555-1234'], source: 'body' }
      ];
      const { deduplicated } = deduplicateContacts(contacts);
      const john = deduplicated.find(c => c.email === 'john@example.com');
      expect(john.displayName).toBe('John');
      expect(john.phoneNumbers).toEqual(['555-1234']);
    });
  });

  describe('calculateConfidence', () => {
    it('should return high confidence for complete contact', () => {
      const contact = {
        email: 'john@example.com',
        displayName: 'John Doe',
        firstName: 'John',
        lastName: 'Doe',
        phoneNumbers: ['555-1234'],
        linkedInUrls: ['https://linkedin.com/in/john'],
        companyName: 'Example Corp',
        jobTitle: 'CEO',
        source: 'metadata'
      };
      expect(calculateConfidence(contact)).toBe('high');
    });

    it('should return low confidence for minimal contact', () => {
      const contact = {
        email: 'unknown@example.com',
        source: 'body'
      };
      expect(calculateConfidence(contact)).toBe('low');
    });
  });

  describe('normalizeContact', () => {
    it('should normalize contact data', () => {
      const contact = {
        email: 'JOHN@EXAMPLE.COM',
        displayName: '  John Doe  ',
        phoneNumber: '555-1234'
      };
      const normalized = normalizeContact(contact);
      expect(normalized.email).toBe('john@example.com');
      expect(normalized.displayName).toBe('John Doe');
      expect(normalized.phoneNumbers).toEqual(['555-1234']);
    });

    it('should calculate confidence if not provided', () => {
      const contact = {
        email: 'john@example.com',
        displayName: 'John Doe'
      };
      const normalized = normalizeContact(contact);
      expect(normalized.extractionConfidence).toBeDefined();
    });
  });
});

describe('CSV Generator', () => {
  describe('escapeCSVField', () => {
    it('should escape fields with commas', () => {
      const field = 'Doe, John';
      const escaped = escapeCSVField(field);
      expect(escaped).toBe('"Doe, John"');
    });

    it('should escape fields with quotes', () => {
      const field = 'Say "Hello"';
      const escaped = escapeCSVField(field);
      expect(escaped).toContain('""');
    });

    it('should return empty string for null/undefined', () => {
      expect(escapeCSVField(null)).toBe('');
      expect(escapeCSVField(undefined)).toBe('');
    });
  });

  describe('arrayToCSVString', () => {
    it('should join array elements', () => {
      const arr = ['555-1234', '555-5678'];
      const result = arrayToCSVString(arr);
      expect(result).toBe('555-1234;555-5678');
    });

    it('should return empty string for empty array', () => {
      expect(arrayToCSVString([])).toBe('');
    });
  });

  describe('generateCSV', () => {
    it('should generate CSV with headers', () => {
      const contacts = [
        {
          email: 'john@example.com',
          displayName: 'John Doe',
          firstName: 'John',
          lastName: 'Doe',
          phoneNumbers: [],
          linkedInUrls: [],
          companyName: null,
          jobTitle: null,
          source: 'metadata',
          isInOutlook: false,
          firstSeenDate: '2025-01-01',
          extractionConfidence: 'high'
        }
      ];
      const csv = generateCSV(contacts);
      expect(csv).toContain('email,displayName,firstName,lastName');
      expect(csv).toContain('john@example.com');
    });

    it('should handle multiple contacts', () => {
      const contacts = [
        { email: 'john@example.com', displayName: 'John', phoneNumbers: [], linkedInUrls: [] },
        { email: 'jane@example.com', displayName: 'Jane', phoneNumbers: [], linkedInUrls: [] }
      ];
      const csv = generateCSV(contacts);
      const lines = csv.split('\n');
      expect(lines.length).toBe(3); // Header + 2 contacts
    });
  });
});

describe('Email Contact Extractor', () => {
  describe('extractContactsFromEmail', () => {
    it('should extract contacts from email metadata', () => {
      const email = {
        from: {
          emailAddress: { name: 'John Doe', address: 'john@example.com' }
        },
        toRecipients: [
          { emailAddress: { name: 'Jane Smith', address: 'jane@example.com' } }
        ],
        ccRecipients: [],
        receivedDateTime: '2025-01-01T12:00:00Z',
        body: { content: '<p>Simple email</p>' }
      };

      const contacts = extractContactsFromEmail(email, false);
      expect(contacts.length).toBe(2);
      expect(contacts.some(c => c.email === 'john@example.com')).toBe(true);
      expect(contacts.some(c => c.email === 'jane@example.com')).toBe(true);
    });

    it('should extract contacts from email body when enabled', () => {
      const email = {
        from: {
          emailAddress: { name: 'Sender', address: 'sender@example.com' }
        },
        toRecipients: [],
        ccRecipients: [],
        receivedDateTime: '2025-01-01T12:00:00Z',
        body: {
          content: '<p>Contact Bob at bob@company.com</p><p>Best regards,<br>Alice<br>CEO, Example Corp<br>alice@example.com<br>555-1234</p>'
        }
      };

      const contacts = extractContactsFromEmail(email, true);
      expect(contacts.length).toBeGreaterThan(1);
    });

    it('should not duplicate contacts from metadata and body', () => {
      const email = {
        from: {
          emailAddress: { name: 'John Doe', address: 'john@example.com' }
        },
        toRecipients: [],
        ccRecipients: [],
        receivedDateTime: '2025-01-01T12:00:00Z',
        body: {
          content: '<p>My email is john@example.com</p>'
        }
      };

      const contacts = extractContactsFromEmail(email, true);
      const johnContacts = contacts.filter(c => c.email === 'john@example.com');
      expect(johnContacts.length).toBe(1); // Should only appear once
    });
  });
});
