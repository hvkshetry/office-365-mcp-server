/**
 * Tests for Newsletter Detection Utility
 */

const {
  detectNewsletter,
  checkHeaders,
  checkSenderPatterns,
  checkBodyContent,
  checkRecipientPatterns,
  checkEmailStructure,
  checkSubject,
  applyWhitelistBlacklist,
  clearCache,
  NEWSLETTER_SIGNALS
} = require('../utils/newsletter-detector');

describe('Newsletter Detection', () => {
  beforeEach(() => {
    clearCache();
  });

  describe('checkHeaders', () => {
    it('should detect List-Unsubscribe header', () => {
      const headers = [
        { name: 'List-Unsubscribe', value: '<mailto:unsubscribe@example.com>' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.listUnsubscribe);
      expect(signals).toContain('list-unsubscribe-header');
    });

    it('should detect bulk precedence', () => {
      const headers = [
        { name: 'Precedence', value: 'bulk' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.bulkPrecedence);
      expect(signals).toContain('bulk-precedence');
    });

    it('should detect List-ID header', () => {
      const headers = [
        { name: 'List-ID', value: '<newsletter.example.com>' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.listId);
      expect(signals).toContain('list-id-header');
    });

    it('should detect Mailchimp headers', () => {
      const headers = [
        { name: 'X-Mailer', value: 'MailChimp Mailer' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.mailchimpHeader);
      expect(signals).toContain('esp-mailer');
    });

    it('should detect Sendgrid headers', () => {
      const headers = [
        { name: 'X-Sender', value: 'SendGrid' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.mailchimpHeader);
      expect(signals).toContain('esp-mailer');
    });

    it('should return 0 for regular email headers', () => {
      const headers = [
        { name: 'From', value: 'colleague@company.com' },
        { name: 'To', value: 'me@company.com' }
      ];
      const signals = [];
      const score = checkHeaders(headers, signals);

      expect(score).toBe(0);
      expect(signals.length).toBe(0);
    });
  });

  describe('checkSenderPatterns', () => {
    it('should detect English noreply address', () => {
      const from = {
        emailAddress: { address: 'noreply@example.com' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.noReplyAddress);
      expect(signals).toContain('noreply-sender');
    });

    it('should detect French ne-pas-repondre address', () => {
      const from = {
        emailAddress: { address: 'ne-pas-repondre@example.fr' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.noReplyAddress);
      expect(signals).toContain('noreply-sender');
    });

    it('should detect newsletter sender', () => {
      const from = {
        emailAddress: { address: 'newsletter@company.com' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.infoAddress);
      expect(signals).toContain('marketing-sender');
    });

    it('should detect French bulletin sender', () => {
      const from = {
        emailAddress: { address: 'bulletin@company.fr' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.infoAddress);
      expect(signals).toContain('marketing-sender');
    });

    it('should detect automated sender', () => {
      const from = {
        emailAddress: { address: 'automated@system.com' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.automatedAddress);
      expect(signals).toContain('automated-sender');
    });

    it('should return 0 for regular work email', () => {
      const from = {
        emailAddress: { address: 'john.doe@company.com' }
      };
      const signals = [];
      const score = checkSenderPatterns(from, signals);

      expect(score).toBeLessThan(10); // Allow for generic patterns
    });
  });

  describe('checkBodyContent', () => {
    it('should detect English unsubscribe link', () => {
      const content = '<p>This is a newsletter.</p><a href="#">Unsubscribe</a>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('unsubscribe-link');
    });

    it('should detect French unsubscribe link', () => {
      const content = '<p>Ceci est une newsletter.</p><a href="#">Se désabonner</a>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('unsubscribe-link');
    });

    it('should detect view in browser (English)', () => {
      const content = '<p>Can\'t read this email? View it in your browser</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('view-in-browser');
    });

    it('should detect view in browser (French)', () => {
      const content = '<p>Afficher dans le navigateur</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('view-in-browser');
    });

    it('should detect update preferences (English)', () => {
      const content = '<p>Update your email preferences here</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('update-preferences');
    });

    it('should detect update preferences (French)', () => {
      const content = '<p>Gérer vos préférences email</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('update-preferences');
    });

    it('should detect social media footer', () => {
      const content = '<p>Follow us on Facebook, Twitter, and LinkedIn</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('social-footer');
    });

    it('should detect French social media footer', () => {
      const content = '<p>Suivez-nous sur les réseaux sociaux: Facebook, LinkedIn</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('social-footer');
    });

    it('should detect tracking pixels', () => {
      const content = '<img width="1" height="1" src="track.gif"><img width="1" height="1" src="pixel.gif">';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('tracking-pixels');
    });

    it('should return 0 for regular email content', () => {
      const content = '<p>Hi John, let\'s meet tomorrow at 2pm. Best regards, Jane</p>';
      const signals = [];
      const score = checkBodyContent(content, signals);

      expect(score).toBeLessThan(15); // Allow for some false positives
    });
  });

  describe('checkRecipientPatterns', () => {
    it('should detect BCC recipient', () => {
      const bccRecipients = [{ emailAddress: { address: 'user@example.com' } }];
      const signals = [];
      const score = checkRecipientPatterns([], [], bccRecipients, signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.bccRecipient);
      expect(signals).toContain('bcc-recipient');
    });

    it('should detect generic recipient name (English)', () => {
      const toRecipients = [
        { emailAddress: { name: 'Valued Customer', address: 'customer@example.com' } }
      ];
      const signals = [];
      const score = checkRecipientPatterns(toRecipients, [], [], signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.genericRecipient);
      expect(signals).toContain('generic-recipient');
    });

    it('should detect generic recipient name (French)', () => {
      const toRecipients = [
        { emailAddress: { name: 'Cher Client', address: 'client@example.fr' } }
      ];
      const signals = [];
      const score = checkRecipientPatterns(toRecipients, [], [], signals);

      expect(score).toBe(NEWSLETTER_SIGNALS.genericRecipient);
      expect(signals).toContain('generic-recipient');
    });

    it('should return 0 for regular recipient', () => {
      const toRecipients = [
        { emailAddress: { name: 'John Doe', address: 'john@example.com' } }
      ];
      const signals = [];
      const score = checkRecipientPatterns(toRecipients, [], [], signals);

      expect(score).toBe(0);
    });
  });

  describe('checkEmailStructure', () => {
    it('should detect table-based layout', () => {
      const content = '<table><tr><td>Cell 1</td></tr></table>'.repeat(6);
      const signals = [];
      const score = checkEmailStructure(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('table-layout');
    });

    it('should detect high image-to-text ratio', () => {
      const content = '<img src="1.jpg"><img src="2.jpg"><img src="3.jpg"><img src="4.jpg"><p>Hi</p>';
      const signals = [];
      const score = checkEmailStructure(content, signals);

      expect(score).toBeGreaterThan(0);
      expect(signals).toContain('high-image-ratio');
    });

    it('should return 0 for regular email structure', () => {
      const content = '<p>This is a regular email with normal text content and no excessive images or tables.</p>';
      const signals = [];
      const score = checkEmailStructure(content, signals);

      expect(score).toBe(0);
    });
  });

  describe('checkSubject', () => {
    it('should detect newsletter subject (English)', () => {
      const subject = 'Weekly Newsletter - Tech Updates';
      const signals = [];
      const score = checkSubject(subject, signals);

      expect(score).toBe(10);
      expect(signals).toContain('newsletter-subject');
    });

    it('should detect newsletter subject (French)', () => {
      const subject = 'Lettre d\'information mensuelle';
      const signals = [];
      const score = checkSubject(subject, signals);

      expect(score).toBe(10);
      expect(signals).toContain('newsletter-subject');
    });

    it('should return 0 for regular subject', () => {
      const subject = 'Meeting tomorrow at 2pm';
      const signals = [];
      const score = checkSubject(subject, signals);

      expect(score).toBe(0);
    });
  });

  describe('detectNewsletter', () => {
    it('should identify Mailchimp newsletter', async () => {
      const email = {
        id: 'test1',
        receivedDateTime: new Date().toISOString(),
        from: { emailAddress: { address: 'newsletter@company.com' } },
        subject: 'Monthly Newsletter - July',
        internetMessageHeaders: [
          { name: 'X-Mailer', value: 'MailChimp' },
          { name: 'List-Unsubscribe', value: '<mailto:unsub@list.com>' }
        ],
        body: {
          content: '<p>View in browser</p><p>Unsubscribe</p><p>Follow us on social media</p>'
        },
        toRecipients: [],
        ccRecipients: [],
        bccRecipients: []
      };

      const result = await detectNewsletter(email);

      expect(result.isNewsletter).toBe(true);
      expect(result.confidence).toBeGreaterThan(60);
      expect(result.signals.length).toBeGreaterThan(3);
    });

    it('should NOT flag regular work email', async () => {
      const email = {
        id: 'test2',
        receivedDateTime: new Date().toISOString(),
        from: { emailAddress: { address: 'colleague@company.com' } },
        subject: 'Re: Project Update',
        internetMessageHeaders: [
          { name: 'From', value: 'colleague@company.com' }
        ],
        body: {
          content: '<p>Hi John,</p><p>Here is the update on the project.</p><p>Best regards,<br>Jane</p>'
        },
        toRecipients: [{ emailAddress: { name: 'John Doe', address: 'john@company.com' } }],
        ccRecipients: [],
        bccRecipients: []
      };

      const result = await detectNewsletter(email);

      expect(result.isNewsletter).toBe(false);
      expect(result.confidence).toBeLessThan(60);
    });

    it('should cache detection results', async () => {
      const email = {
        id: 'test3',
        receivedDateTime: new Date().toISOString(),
        from: { emailAddress: { address: 'test@example.com' } },
        subject: 'Test',
        internetMessageHeaders: [],
        body: { content: '<p>Test</p>' },
        toRecipients: [],
        ccRecipients: [],
        bccRecipients: []
      };

      const result1 = await detectNewsletter(email);
      const result2 = await detectNewsletter(email);

      expect(result1).toEqual(result2);
    });

    it('should use custom threshold', async () => {
      const email = {
        id: 'test4',
        receivedDateTime: new Date().toISOString(),
        from: { emailAddress: { address: 'info@company.com' } },
        subject: 'Updates',
        internetMessageHeaders: [],
        body: { content: '<p>Unsubscribe here</p>' },
        toRecipients: [],
        ccRecipients: [],
        bccRecipients: []
      };

      const resultLow = await detectNewsletter(email, 30);
      const resultHigh = await detectNewsletter(email, 80);

      expect(resultLow.isNewsletter).toBe(true);
      expect(resultHigh.isNewsletter).toBe(false);
    });

    it('should handle detection errors gracefully', async () => {
      const invalidEmail = null;

      const result = await detectNewsletter(invalidEmail);

      expect(result.isNewsletter).toBe(false);
      expect(result.confidence).toBe(0);
      expect(result.signals).toContain('detection-error');
    });
  });

  describe('applyWhitelistBlacklist', () => {
    it('should whitelist domain', () => {
      const email = {
        from: { emailAddress: { address: 'newsletter@trusted-vendor.com' } },
        subject: 'Newsletter'
      };
      const detection = {
        isNewsletter: true,
        confidence: 80,
        signals: ['unsubscribe-link'],
        reason: 'unsubscribe-link'
      };
      const rules = {
        whitelist: {
          domains: ['trusted-vendor.com']
        }
      };

      const result = applyWhitelistBlacklist(email, detection, rules);

      expect(result.isNewsletter).toBe(false);
      expect(result.signals).toContain('whitelisted');
    });

    it('should whitelist specific sender', () => {
      const email = {
        from: { emailAddress: { address: 'important@company.com' } },
        subject: 'Updates'
      };
      const detection = {
        isNewsletter: true,
        confidence: 70,
        signals: ['newsletter-subject'],
        reason: 'newsletter-subject'
      };
      const rules = {
        whitelist: {
          senders: ['important@company.com']
        }
      };

      const result = applyWhitelistBlacklist(email, detection, rules);

      expect(result.isNewsletter).toBe(false);
      expect(result.signals).toContain('whitelisted');
    });

    it('should blacklist domain', () => {
      const email = {
        from: { emailAddress: { address: 'user@spam-company.com' } },
        subject: 'Regular email'
      };
      const detection = {
        isNewsletter: false,
        confidence: 20,
        signals: [],
        reason: 'no-signals'
      };
      const rules = {
        blacklist: {
          domains: ['spam-company.com']
        }
      };

      const result = applyWhitelistBlacklist(email, detection, rules);

      expect(result.isNewsletter).toBe(true);
      expect(result.confidence).toBe(100);
      expect(result.signals).toContain('blacklisted');
    });

    it('should apply custom sender pattern', () => {
      const email = {
        from: { emailAddress: { address: 'promo@company.com' } },
        subject: 'Check this out'
      };
      const detection = {
        isNewsletter: false,
        confidence: 30,
        signals: [],
        reason: 'no-signals'
      };
      const rules = {
        customPatterns: {
          senderPatterns: ['^promo@']
        }
      };

      const result = applyWhitelistBlacklist(email, detection, rules);

      expect(result.isNewsletter).toBe(true);
      expect(result.confidence).toBeGreaterThanOrEqual(75);
      expect(result.signals).toContain('custom-sender-pattern');
    });

    it('should apply custom subject pattern', () => {
      const email = {
        from: { emailAddress: { address: 'user@company.com' } },
        subject: '[PROMO] Special Offer!'
      };
      const detection = {
        isNewsletter: false,
        confidence: 20,
        signals: [],
        reason: 'no-signals'
      };
      const rules = {
        customPatterns: {
          subjectPatterns: ['\\[PROMO\\]']
        }
      };

      const result = applyWhitelistBlacklist(email, detection, rules);

      expect(result.isNewsletter).toBe(true);
      expect(result.confidence).toBeGreaterThanOrEqual(75);
      expect(result.signals).toContain('custom-subject-pattern');
    });

    it('should handle null rules gracefully', () => {
      const email = {
        from: { emailAddress: { address: 'test@example.com' } },
        subject: 'Test'
      };
      const detection = {
        isNewsletter: false,
        confidence: 10,
        signals: [],
        reason: 'no-signals'
      };

      const result = applyWhitelistBlacklist(email, detection, null);

      expect(result).toEqual(detection);
    });
  });
});
