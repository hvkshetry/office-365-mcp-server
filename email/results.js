// Mail has one MCP tool definition for all operations. Once it declares an
// outputSchema, every successful operation must provide structured content.
const mailOutputSchema = {
  type: 'object',
  properties: {
    text: { type: 'string', description: 'Complete readable result, including IDs and coverage notes' },
    emails: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          subject: { type: 'string' },
          from: { type: 'string' },
          receivedDateTime: { type: ['string', 'null'] },
          hasAttachments: { type: 'boolean' }
        },
        required: ['id', 'subject', 'from', 'receivedDateTime', 'hasAttachments']
      }
    },
    method: { type: 'string' },
    returned: { type: 'integer' },
    pageLimit: { type: 'number' },
    hasMore: { type: 'boolean' },
    limited: { type: 'boolean' },
    serviceLimit: { type: ['integer', 'null'] },
    note: { type: 'string' }
  },
  required: ['text']
};

function withMailStructuredContent(result) {
  const { content = [], structuredContent = {}, isError, ...metadata } = result;
  const text = content.filter(item => item.type === 'text').map(item => item.text).join('\n');
  // Keep the readable text blocks for existing callers. Structured-first clients
  // need the entire result too, including metadata such as draftId and sent.
  return { ...result, structuredContent: { ...metadata, ...structuredContent, text } };
}

module.exports = { mailOutputSchema, withMailStructuredContent };
