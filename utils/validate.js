/**
 * Input validation helpers for Graph API parameters.
 * Defense-in-depth: reject path traversal and injection chars in IDs
 * before they reach URL construction.
 */

const DANGEROUS_PATTERNS = /[\/\\]|\.\./;

/**
 * Validates that an ID parameter is safe to interpolate into a URL path.
 * Graph API IDs are typically alphanumeric with hyphens, underscores, or base64 chars.
 * @param {string} id - The ID value to validate
 * @param {string} paramName - Name of the parameter (for error messages)
 * @returns {string} The validated ID
 * @throws {Error} If the ID contains dangerous characters
 */
function validateId(id, paramName = 'id') {
  if (!id || typeof id !== 'string') {
    throw new Error(`Missing or invalid ${paramName}`);
  }
  if (DANGEROUS_PATTERNS.test(id)) {
    throw new Error(`Invalid ${paramName}: contains disallowed characters`);
  }
  return id;
}

module.exports = { validateId };
