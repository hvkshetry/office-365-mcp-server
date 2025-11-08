/**
 * Newsletter Detection Utility (Bilingual: French & English)
 * Multi-factor analysis to identify and filter newsletter/marketing emails
 */

// Signal scoring weights
const NEWSLETTER_SIGNALS = {
  // Email Headers & Metadata (High confidence)
  listUnsubscribe: 25,
  bulkPrecedence: 20,
  listId: 20,
  mailchimpHeader: 25,

  // Sender Patterns (Medium confidence)
  noReplyAddress: 15,
  infoAddress: 10,
  automatedAddress: 12,

  // Content Patterns (Medium confidence)
  unsubscribeLink: 20,
  viewInBrowser: 15,
  updatePreferences: 12,
  manageSubscription: 10,

  // Structure Patterns (Low-Medium confidence)
  highImageToTextRatio: 10,
  multipleTrackingPixels: 8,
  tableBasedLayout: 5,
  socialMediaFooter: 8,

  // Recipient Patterns (Medium confidence)
  bccRecipient: 15,
  genericRecipient: 10
};

// Detection cache
const detectionCache = new Map();
const CACHE_MAX_SIZE = 1000;

/**
 * Main newsletter detection function with caching
 * @param {Object} email - Email object from Graph API
 * @param {number} threshold - Confidence threshold (default: 60)
 * @returns {Promise<Object>} { isNewsletter, confidence, signals, reason }
 */
async function detectNewsletter(email, threshold = 60) {
  // Handle null/invalid email
  if (!email || !email.id) {
    return {
      isNewsletter: false,
      confidence: 0,
      signals: ['detection-error'],
      reason: 'detection-error'
    };
  }

  const cacheKey = `${email.id}-${email.receivedDateTime}-${threshold}`;

  if (detectionCache.has(cacheKey)) {
    return detectionCache.get(cacheKey);
  }

  try {
    const signals = [];
    let score = 0;

    // Check headers
    score += checkHeaders(email.internetMessageHeaders, signals);

    // Check sender patterns
    score += checkSenderPatterns(email.from, signals);

    // Check body content
    score += checkBodyContent(email.body?.content, signals);

    // Check recipient patterns
    score += checkRecipientPatterns(
      email.toRecipients,
      email.ccRecipients,
      email.bccRecipients,
      signals
    );

    // Check email structure
    score += checkEmailStructure(email.body?.content, signals);

    // Check subject line
    score += checkSubject(email.subject, signals);

    const result = {
      isNewsletter: score >= threshold,
      confidence: Math.min(score, 100),
      signals: signals,
      reason: signals.slice(0, 3).join(', ') || 'no-signals'
    };

    // Cache the result
    if (detectionCache.size >= CACHE_MAX_SIZE) {
      const firstKey = detectionCache.keys().next().value;
      detectionCache.delete(firstKey);
    }
    detectionCache.set(cacheKey, result);

    return result;

  } catch (error) {
    console.error(`Newsletter detection failed for email ${email.id}:`, error);
    // Fail-safe: treat as non-newsletter if detection fails
    return {
      isNewsletter: false,
      confidence: 0,
      signals: ['detection-error'],
      reason: 'detection-error'
    };
  }
}

/**
 * Check email headers for newsletter indicators
 * @param {Array} headers - Internet message headers
 * @param {Array} signals - Signal array to populate
 * @returns {number} - Score contribution
 */
function checkHeaders(headers, signals) {
  let score = 0;

  if (!headers || !Array.isArray(headers)) return score;

  const headerMap = {};
  headers.forEach(h => {
    if (h.name && h.value) {
      headerMap[h.name.toLowerCase()] = h.value.toLowerCase();
    }
  });

  // List-Unsubscribe header
  if (headerMap['list-unsubscribe']) {
    signals.push('list-unsubscribe-header');
    score += NEWSLETTER_SIGNALS.listUnsubscribe;
  }

  // Precedence: bulk/list
  const precedence = headerMap['precedence'] || '';
  if (precedence.includes('bulk') || precedence.includes('list')) {
    signals.push('bulk-precedence');
    score += NEWSLETTER_SIGNALS.bulkPrecedence;
  }

  // List-ID header
  if (headerMap['list-id']) {
    signals.push('list-id-header');
    score += NEWSLETTER_SIGNALS.listId;
  }

  // ESP headers (Mailchimp, Sendgrid, Campaign Monitor, etc.)
  const mailer = headerMap['x-mailer'] || headerMap['x-sender'] || headerMap['x-campaign'] || '';
  const espPatterns = /mailchimp|sendgrid|constant contact|campaign monitor|mailgun|postmark|amazonses|sendinblue|brevo|mailjet/i;
  if (espPatterns.test(mailer)) {
    signals.push('esp-mailer');
    score += NEWSLETTER_SIGNALS.mailchimpHeader;
  }

  return score;
}

/**
 * Check sender email patterns (French & English)
 * @param {Object} from - From email address object
 * @param {Array} signals - Signal array to populate
 * @returns {number} - Score contribution
 */
function checkSenderPatterns(from, signals) {
  let score = 0;

  if (!from?.emailAddress?.address) return score;

  const email = from.emailAddress.address.toLowerCase();

  // No-reply patterns (English & French)
  if (/^(no-?reply|do-?not-?reply|noreply|ne-?pas-?repondre|nepasrepondre)@/i.test(email)) {
    signals.push('noreply-sender');
    score += NEWSLETTER_SIGNALS.noReplyAddress;
  }

  // Info/Newsletter/Marketing patterns (English & French)
  if (/^(info|newsletter|marketing|news|updates|notifications|lettre|actualit(é|e)s?|communication|diffusion|bulletin)@/i.test(email)) {
    signals.push('marketing-sender');
    score += NEWSLETTER_SIGNALS.infoAddress;
  }

  // Automated patterns (English & French)
  if (/^(automated?|auto|notification|alerts?|system|automatique|alerte)@/i.test(email)) {
    signals.push('automated-sender');
    score += NEWSLETTER_SIGNALS.automatedAddress;
  }

  // Generic French patterns
  if (/^(contact|service|support|accueil|equipe)@/i.test(email)) {
    signals.push('generic-french-sender');
    score += 8;
  }

  return score;
}

/**
 * Check email body content for newsletter patterns (French & English)
 * @param {string} htmlContent - HTML body content
 * @param {Array} signals - Signal array to populate
 * @returns {number} - Score contribution
 */
function checkBodyContent(htmlContent, signals) {
  let score = 0;

  if (!htmlContent) return score;

  const lowerContent = htmlContent.toLowerCase();

  // Unsubscribe link (English & French)
  if (/unsubscribe|se d(é|e)sabonner|d(é|e)sabonnement|d(é|e)sinscription|ne plus recevoir/i.test(lowerContent)) {
    signals.push('unsubscribe-link');
    score += NEWSLETTER_SIGNALS.unsubscribeLink;
  }

  // View in browser (English & French)
  if (/(view|read|voir|lire|afficher|consulter).{0,50}(this|email|message|ce|cet|cette|it).{0,50}(in|on|dans|sur).{0,20}(your |le |votre |ton |ta )?(navigat(eur|or)|browser)|afficher dans le navigateur|voir en ligne/i.test(lowerContent)) {
    signals.push('view-in-browser');
    score += NEWSLETTER_SIGNALS.viewInBrowser;
  }

  // Update preferences (English & French)
  if (/(update|manage|change|modifier|g(é|e)rer|mettre (à|a) jour).{0,50}(your |vos |mes |tes )?((e-?mail|courriel) )?pr(é|e)f(é|e)rences/i.test(lowerContent)) {
    signals.push('update-preferences');
    score += NEWSLETTER_SIGNALS.updatePreferences;
  }

  // Manage subscription (English & French)
  if (/(manage|modify|g(é|e)rer|modifier).{0,30}(your |votre |mon |ton )?(subscription|abonnement|inscription)/i.test(lowerContent)) {
    signals.push('manage-subscription');
    score += NEWSLETTER_SIGNALS.manageSubscription;
  }

  // Newsletter specific phrases (French)
  if (/(cet?|ce) (e-?mail|courriel|message) (vous |t')?(a (é|e)t(é|e)|est) envoy(é|e)|vous recevez ce (message|mail|courriel)|lettre d'information|infolettre|bulletin d'information/i.test(lowerContent)) {
    signals.push('french-newsletter-phrase');
    score += 12;
  }

  // Newsletter specific phrases (English)
  if (/(this|the) (email|message|newsletter) (was|has been) sent|you (are receiving|received) this (email|message)|email newsletter|mailing list/i.test(lowerContent)) {
    signals.push('english-newsletter-phrase');
    score += 12;
  }

  // Tracking pixels (1x1 images)
  const trackingPixels = (htmlContent.match(/width\s*=\s*["']?1["']?\s+height\s*=\s*["']?1["']?|height\s*=\s*["']?1["']?\s+width\s*=\s*["']?1["']?/gi) || []).length;
  if (trackingPixels >= 2) {
    signals.push('tracking-pixels');
    score += NEWSLETTER_SIGNALS.multipleTrackingPixels;
  }

  // Social media footer (English & French)
  if (/(follow us|connect with us|find us on|suivez[ -]nous|retrouvez[ -]nous|rejoignez[ -]nous|suivre|suivez).{0,100}(facebook|twitter|linkedin|instagram|social|r(é|e)seaux sociaux)/i.test(lowerContent)) {
    signals.push('social-footer');
    score += NEWSLETTER_SIGNALS.socialMediaFooter;
  }

  // Privacy/Legal footer phrases (French)
  if (/politique de confidentialit(é|e)|protection des donn(é|e)es|mentions l(é|e)gales|conform(é|e)ment (à|a) la loi|cnil|rgpd/i.test(lowerContent)) {
    signals.push('privacy-footer');
    score += 5;
  }

  // Privacy/Legal footer phrases (English)
  if (/privacy policy|data protection|legal notice|terms (of service|and conditions)|gdpr|unsubscribe policy/i.test(lowerContent)) {
    signals.push('privacy-footer-en');
    score += 5;
  }

  return score;
}

/**
 * Check recipient patterns (French & English)
 * @param {Array} toRecipients - To recipients
 * @param {Array} ccRecipients - CC recipients
 * @param {Array} bccRecipients - BCC recipients
 * @param {Array} signals - Signal array to populate
 * @returns {number} - Score contribution
 */
function checkRecipientPatterns(toRecipients, ccRecipients, bccRecipients, signals) {
  let score = 0;

  // Check if user is in BCC (bulk sending indicator)
  if (bccRecipients && bccRecipients.length > 0) {
    signals.push('bcc-recipient');
    score += NEWSLETTER_SIGNALS.bccRecipient;
  }

  // Check for generic recipient names (English & French)
  const toNames = toRecipients?.map(r => r.emailAddress?.name?.toLowerCase() || '') || [];
  const genericPatterns = /valued customer|dear subscriber|dear user|dear member|newsletter subscriber|cher client|cher abonn(é|e)|cher membre|cher(e)? utilisateur|bonjour (à|a) tous|madame,? monsieur|cher lecteur|ami(e)? lecteur/i;

  if (toNames.some(name => genericPatterns.test(name))) {
    signals.push('generic-recipient');
    score += NEWSLETTER_SIGNALS.genericRecipient;
  }

  return score;
}

/**
 * Analyze email HTML structure
 * @param {string} htmlContent - HTML body content
 * @param {Array} signals - Signal array to populate
 * @param {number} [tableLayoutThreshold=5] - Threshold for table-based layout detection
 * @returns {number} - Score contribution
 */
function checkEmailStructure(htmlContent, signals, tableLayoutThreshold = 5) {
  let score = 0;

  if (!htmlContent) return score;

  // Count tables (newsletters often use table-based layouts)
  const tableCount = (htmlContent.match(/<table/gi) || []).length;
  if (tableCount > tableLayoutThreshold) {
    signals.push('table-layout');
    score += NEWSLETTER_SIGNALS.tableBasedLayout;
  }

  // Image to text ratio
  const imageCount = (htmlContent.match(/<img/gi) || []).length;
  const textLength = htmlContent.replace(/<[^>]*>/g, '').trim().length;

  if (textLength > 0 && imageCount > 0) {
    const imageRatio = imageCount / Math.max(textLength / 100, 1);

    if (imageRatio > 0.7) {
      signals.push('high-image-ratio');
      score += NEWSLETTER_SIGNALS.highImageToTextRatio;
    }
  }

  return score;
}

/**
 * Check subject line for newsletter patterns
 * @param {string} subject - Email subject
 * @param {Array} signals - Signal array to populate
 * @returns {number} - Score contribution
 */
function checkSubject(subject, signals) {
  let score = 0;

  if (!subject) return score;

  const lowerSubject = subject.toLowerCase();

  // Newsletter indicators in subject (English & French)
  if (/(newsletter|bulletin|digest|roundup|weekly|monthly|update|actualit(é|e)s?|lettre d'information|hebdomadaire|mensuel|r(é|e)capitulatif)/i.test(lowerSubject)) {
    signals.push('newsletter-subject');
    score += 10;
  }

  return score;
}

/**
 * Apply whitelist/blacklist rules to override detection
 * @param {Object} email - Email object
 * @param {Object} detection - Detection result
 * @param {Object} rules - Rules from newsletter-rules.json
 * @returns {Object} - Modified detection result
 */
function applyWhitelistBlacklist(email, detection, rules) {
  if (!rules || !email?.from?.emailAddress?.address) {
    return detection;
  }

  const senderEmail = email.from.emailAddress.address.toLowerCase();
  const senderDomain = senderEmail.split('@')[1] || '';

  // Check whitelist (never filter)
  if (rules.whitelist) {
    const whitelistedDomains = (rules.whitelist.domains || []).map(d => d.toLowerCase());
    const whitelistedSenders = (rules.whitelist.senders || []).map(s => s.toLowerCase());

    if (whitelistedDomains.includes(senderDomain) || whitelistedSenders.includes(senderEmail)) {
      return {
        ...detection,
        isNewsletter: false,
        signals: [...detection.signals, 'whitelisted'],
        reason: 'whitelisted'
      };
    }
  }

  // Check blacklist (always filter)
  if (rules.blacklist) {
    const blacklistedDomains = (rules.blacklist.domains || []).map(d => d.toLowerCase());
    const blacklistedSenders = (rules.blacklist.senders || []).map(s => s.toLowerCase());

    if (blacklistedDomains.includes(senderDomain) || blacklistedSenders.includes(senderEmail)) {
      return {
        ...detection,
        isNewsletter: true,
        confidence: 100,
        signals: [...detection.signals, 'blacklisted'],
        reason: 'blacklisted'
      };
    }
  }

  // Check custom patterns
  if (rules.customPatterns) {
    const senderPatterns = rules.customPatterns.senderPatterns || [];
    const subjectPatterns = rules.customPatterns.subjectPatterns || [];

    // Check sender patterns
    for (const pattern of senderPatterns) {
      try {
        const regex = new RegExp(pattern, 'i');
        if (regex.test(senderEmail)) {
          return {
            ...detection,
            isNewsletter: true,
            confidence: Math.max(detection.confidence, 75),
            signals: [...detection.signals, 'custom-sender-pattern'],
            reason: 'custom-sender-pattern'
          };
        }
      } catch (e) {
        console.error(`Invalid sender pattern: ${pattern}`, e);
      }
    }

    // Check subject patterns
    const subject = email.subject || '';
    for (const pattern of subjectPatterns) {
      try {
        const regex = new RegExp(pattern, 'i');
        if (regex.test(subject)) {
          return {
            ...detection,
            isNewsletter: true,
            confidence: Math.max(detection.confidence, 75),
            signals: [...detection.signals, 'custom-subject-pattern'],
            reason: 'custom-subject-pattern'
          };
        }
      } catch (e) {
        console.error(`Invalid subject pattern: ${pattern}`, e);
      }
    }
  }

  return detection;
}

/**
 * Clear detection cache (for testing or memory management)
 */
function clearCache() {
  detectionCache.clear();
}

module.exports = {
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
};
