/**
 * Structured, read-only source enumeration for procurement reconciliation.
 * All requests use the existing Office delegated account. Cursors are scoped to
 * their originating endpoint; no arbitrary Graph paths or mutation verbs exist.
 */
const { getAccessToken } = require('./auth/token-manager');
const { callGraphAPI } = require('./utils/graph-api');

function identifier(value, name) {
  if (typeof value !== 'string' || !value || /[/\\\x00-\x1f]/.test(value)) {
    throw new Error(`A valid ${name} is required`);
  }
  return encodeURIComponent(value);
}

function requestFor(p) {
  const id = name => identifier(p[name], name);
  let endpoint;
  let query = { $top: 100 };
  const user = () => `users/${id('mailbox')}`;
  const drive = () => `drives/${id('drive_id')}`;
  const item = () => `${drive()}/items/${id('item_id')}`;
  switch (p.operation) {
    case 'sites':
      endpoint = 'sites'; query.search = '*'; break;
    case 'users':
      endpoint = 'users'; query.$select = 'id,mail,userPrincipalName,proxyAddresses'; break;
    case 'drives':
      endpoint = `sites/${id('site_id')}/drives`; break;
    case 'user_drives':
      endpoint = `${user()}/drives`; break;
    case 'folders':
      endpoint = p.folder_id
        ? `${user()}/mailFolders/${id('folder_id')}/childFolders`
        : `${user()}/mailFolders`;
      query.includeHiddenFolders = true;
      break;
    case 'drive_delta':
      endpoint = `${drive()}/root/delta`;
      query = { $select: 'id,name,size,file,folder,parentReference,eTag,cTag,webUrl,lastModifiedDateTime,deleted,remoteItem' };
      break;
    case 'mail_delta':
      endpoint = `${user()}/mailFolders/${id('folder_id')}/messages/delta`;
      query = { $select: 'id,subject,from,receivedDateTime,lastModifiedDateTime,internetMessageId,hasAttachments,isDraft,body' };
      break;
    case 'attachments':
      endpoint = `${user()}/messages/${id('message_id')}/attachments`; break;
    case 'attachment':
      endpoint = `${user()}/messages/${id('message_id')}/attachments/${id('attachment_id')}`;
      query = {}; break;
    case 'file':
      endpoint = item(); query = {}; break;
    case 'file_content':
      endpoint = `${item()}/content`; query = {}; break;
    case 'version_content':
      endpoint = `${item()}/versions/${id('version_id')}/content`; query = {}; break;
    case 'versions':
      endpoint = `${item()}/versions`; break;
    default: throw new Error('Unknown procurement read operation');
  }
  if (p.cursor) {
    const next = new URL(p.cursor);
    if (next.protocol !== 'https:' || next.hostname !== 'graph.microsoft.com' ||
        next.port || next.username || next.password || next.hash ||
        decodeURIComponent(next.pathname) !== decodeURIComponent(`/v1.0/${endpoint}`)) {
      throw new Error('Cursor does not belong to this source endpoint');
    }
    return { endpoint: next.href, query: {} };
  }
  return { endpoint, query };
}

async function handler(params) {
  try {
    const { endpoint, query } = requestFor(params);
    const token = await getAccessToken();
    if (!token) throw new Error('Office authentication unavailable');
    if (['file_content', 'version_content'].includes(params.operation)) {
      const response = await fetch(`https://graph.microsoft.com/v1.0/${endpoint}`, {
        headers: { Authorization: `Bearer ${token}` }, redirect: 'manual',
        signal: AbortSignal.timeout(30000)
      });
      if (response.status !== 302 || !response.headers.get('location')) {
        throw new Error(`Content download location unavailable (${response.status})`);
      }
      const data = { download_url: response.headers.get('location') };
      return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data };
    }
    const data = await callGraphAPI(token, 'GET', endpoint, null, query, {
      Prefer: 'IdType="ImmutableId", odata.maxpagesize=100'
    });
    return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data };
  } catch (error) {
    // Never serialize request URLs: download and delta URLs can contain secrets.
    const code = error.graphError?.code || error.code || 'source_read_failed';
    return { isError: true, content: [{ type: 'text', text: JSON.stringify({
      error: code, status: error.statusCode || null,
      message: error.statusCode ? 'Source access failed; no completeness claim is possible.' : error.message
    }) }] };
  }
}

const tool = {
  name: 'procurement_read',
  description: 'Read structured Microsoft 365 source inventories and changes, preserving pagination. Does not send mail or modify files. Pass the same source IDs with each nextLink/deltaLink cursor.',
  inputSchema: {
    type: 'object', additionalProperties: false, required: ['operation'],
    properties: {
      operation: { type: 'string', enum: ['sites', 'users', 'drives', 'user_drives',
        'folders', 'drive_delta', 'mail_delta', 'attachments', 'attachment', 'file', 'versions',
        'file_content', 'version_content'] },
      ...Object.fromEntries(['site_id', 'drive_id', 'item_id', 'mailbox', 'folder_id',
        'message_id', 'attachment_id', 'version_id', 'cursor'].map(k => [k, { type: 'string' }]))
    }
  }, handler
};
module.exports = { tool, requestFor };
