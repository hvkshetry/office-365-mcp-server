/**
 * TEMPORARY Send As fallback (added 2026-07-23).
 *
 * The primary delegated identity (admin.agent@) holds Exchange Full Access on the
 * persona mailboxes but NOT the "Send As" right, so send-type mail operations for
 * those mailboxes fail with ErrorSendAsDenied until the right is granted
 * (Cinch IT ticket T20260717.0187 follow-up). As a stopgap, send/reply/send_draft
 * for the mailboxes listed in OFFICE_SENDAS_FALLBACK_MAILBOXES use the delegated
 * token stored at OFFICE_SENDAS_TOKEN_PATH instead of the primary token. All other
 * operations (and all other mailboxes) are unaffected.
 *
 * REVERT: remove OFFICE_SENDAS_TOKEN_PATH and OFFICE_SENDAS_FALLBACK_MAILBOXES
 * from the .env and delete the fallback token file. With the env vars unset this
 * module is inert.
 */
const fs = require('fs');
const { refreshAccessToken, needsRefresh } = require('./auto-refresh');

let cachedTokens = null;

function fallbackMailboxes() {
  return (process.env.OFFICE_SENDAS_FALLBACK_MAILBOXES || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);
}

async function getSendAsAccessToken() {
  const tokenPath = process.env.OFFICE_SENDAS_TOKEN_PATH;
  if (!tokenPath || !fs.existsSync(tokenPath)) {
    return null;
  }

  if (cachedTokens && cachedTokens.access_token && !needsRefresh(cachedTokens)) {
    return cachedTokens.access_token;
  }

  let tokens;
  try {
    tokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
  } catch (e) {
    console.error('[SENDAS-FALLBACK] Failed to read fallback token file:', e.message);
    return null;
  }

  if (needsRefresh(tokens)) {
    try {
      tokens = await refreshAccessToken(tokenPath);
    } catch (e) {
      console.error('[SENDAS-FALLBACK] Fallback token refresh failed:', e.message);
      return tokens.access_token || null;
    }
  }

  cachedTokens = tokens;
  return tokens.access_token || null;
}

/**
 * Returns the token to use for a send-type mail operation: the fallback token when
 * the target mailbox is configured for Send As fallback, otherwise the default.
 */
async function resolveSendToken(defaultToken, params) {
  const mailbox = ((params && params.mailbox) || '').toLowerCase();
  if (!mailbox || !fallbackMailboxes().includes(mailbox)) {
    return defaultToken;
  }
  const sendToken = await getSendAsAccessToken();
  if (sendToken) {
    console.error(`[SENDAS-FALLBACK] Using Send As fallback token for ${mailbox}`);
    return sendToken;
  }
  console.error(`[SENDAS-FALLBACK] Fallback configured but token unavailable for ${mailbox}; using primary token`);
  return defaultToken;
}

module.exports = {
  resolveSendToken,
  getSendAsAccessToken,
};
