/**
 * Centralized error helper for MCP tool responses.
 * Pattern adopted from sibling servers (QuickBooks, InvenTree, Atlas CMMS).
 */

const ERROR_HINTS = {
  '401': 'Authentication expired. Run the system tool with operation: "authenticate".',
  '403': 'Permission denied. Check your Microsoft 365 app permissions.',
  '404': 'Resource not found. Verify the ID is correct.',
  '429': 'Rate limit exceeded. Wait a moment and try again.',
  '400': 'Invalid request. Check parameter values.',
  'timeout': 'Request timed out. Try again or reduce the scope.',
  'ECONNREFUSED': 'Connection refused. Check network/VPN.',
};

function getHint(error) {
  const msg = String(error.message || error);
  for (const [key, hint] of Object.entries(ERROR_HINTS)) {
    if (msg.includes(key)) return hint;
  }
  return null;
}

function formatError(error, context = '') {
  const prefix = context ? `[${context}] ` : '';
  const message = error.message || String(error);
  const hint = getHint(error);
  const text = hint
    ? `${prefix}Error: ${message}\nHint: ${hint}`
    : `${prefix}Error: ${message}`;

  console.error(`${prefix}${message}`);
  return {
    isError: true,
    content: [{ type: "text", text }]
  };
}

function safeTool(toolName, handler) {
  return async function(args) {
    try {
      return await handler(args);
    } catch (error) {
      const op = args?.operation || args?.entity || 'unknown';
      return formatError(error, `${toolName}.${op}`);
    }
  };
}

module.exports = { formatError, safeTool, getHint };
