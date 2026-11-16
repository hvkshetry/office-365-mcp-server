/**
 * Document attachments (PDF, Word, Excel, PowerPoint) for Teams channel posts and replies.
 *
 * Unlike images, Graph cannot carry a document inside a channel message. A Teams
 * file attachment is a "reference" attachment pointing at a driveItem in the
 * channel's own SharePoint folder (GET teams/{t}/channels/{c}/filesFolder), which
 * is exactly what the Teams client does when a person attaches a file. So each
 * file is uploaded there first, then referenced from the message:
 *   attachments: [{id: <driveItem eTag GUID>, contentType: "reference", contentUrl, name}]
 *   body:        ... <attachment id="GUID"></attachment> ...
 *
 * Audience: the channel's folder belongs to the team's site, so the file is
 * readable by the team's members, the same people who can read the post. No
 * sharing link is created. The caller's channel fence (policy.enforceTeamsPolicy)
 * runs before this, so uploads only ever land in an allowlisted channel's folder.
 *
 * Callers pass local paths. Because this reads local files on the caller's
 * behalf, a file is accepted only if its bytes are really the document type its
 * extension claims (PDF header, or an OOXML package with that format's main
 * part), and macro-enabled packages are refused. A key or an .env renamed to
 * .pdf or .xlsx is rejected before anything is uploaded.
 */

const fs = require('fs');
const path = require('path');
const { callGraphAPI } = require('../../utils/graph-api');

const MAX_FILES = 5;
const MAX_FILE_BYTES = 25 * 1024 * 1024;
// Graph's simple PUT is for small files; above this an upload session is used.
const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;
// Upload-session chunks must be multiples of 320 KiB.
const CHUNK_BYTES = 320 * 1024 * 12;

const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
const OOXML_MAIN_PART = {
  '.docx': 'word/document.xml',
  '.xlsx': 'xl/workbook.xml',
  '.pptx': 'ppt/presentation.xml'
};
const SUPPORTED = ['.pdf', ...Object.keys(OOXML_MAIN_PART)];

/** Null when `bytes` really is a document of the type `ext` names; else the reason. */
function _documentProblem(ext, bytes) {
  if (ext === '.pdf') {
    return bytes.subarray(0, 5).toString('latin1') === '%PDF-' ? null : 'not a PDF file';
  }
  const main = OOXML_MAIN_PART[ext];
  if (!main) return `unsupported type ${ext || '(none)'}; supported: ${SUPPORTED.join(', ')}`;
  if (!bytes.subarray(0, 4).equals(ZIP_MAGIC)) return `not a ${ext} file`;
  // Part names sit uncompressed in the zip's local and central directory headers.
  const names = bytes.toString('latin1');
  if (!names.includes('[Content_Types].xml') || !names.includes(main)) return `not a ${ext} file`;
  if (names.includes('vbaProject.bin')) return 'macro-enabled documents are not sent';
  return null;
}

/**
 * Validate every file before anything is uploaded.
 * @returns {{files: Array<{path, name, bytes}>}|{error: string}}
 */
function validateFiles(files) {
  if (files === undefined || files === null) return { files: [] };
  if (!Array.isArray(files)) return { error: 'files must be an array of {path, name?}.' };
  if (files.length > MAX_FILES) return { error: `At most ${MAX_FILES} files per message.` };
  const out = [];
  for (const [i, file] of files.entries()) {
    const p = file && file.path;
    if (typeof p !== 'string' || !path.isAbsolute(p)) {
      return { error: `files[${i}].path must be an absolute path to a ${SUPPORTED.join('/')} file.` };
    }
    let stat;
    try {
      stat = fs.statSync(p);
    } catch {
      return { error: `files[${i}]: file not found: ${p}` };
    }
    if (!stat.isFile()) return { error: `files[${i}]: not a regular file: ${p}` };
    if (stat.size === 0) return { error: `files[${i}]: empty file: ${p}` };
    if (stat.size > MAX_FILE_BYTES) {
      return { error: `files[${i}]: ${stat.size} bytes exceeds the ${MAX_FILE_BYTES}-byte limit.` };
    }
    const name = path.basename(file.name || p);
    if (!name || name.startsWith('.') || /[\\/:*?"<>|#%]/.test(name)) {
      return { error: `files[${i}]: unusable file name ${JSON.stringify(name)}` };
    }
    const ext = path.extname(name).toLowerCase();
    if (path.extname(p).toLowerCase() !== ext) {
      return { error: `files[${i}]: name ${name} must keep the file's extension ${path.extname(p)}` };
    }
    const bytes = fs.readFileSync(p);
    const problem = _documentProblem(ext, bytes);
    if (problem) return { error: `files[${i}]: ${problem}: ${p}` };
    out.push({ path: p, name, bytes });
  }
  return { files: out };
}

async function _uploadSession(accessToken, driveId, folderId, file) {
  const session = await callGraphAPI(
    accessToken, 'POST',
    `drives/${driveId}/items/${folderId}:/${encodeURIComponent(file.name)}:/createUploadSession`,
    { item: { '@microsoft.graph.conflictBehavior': 'rename' } }
  );
  const total = file.bytes.length;
  let item = null;
  for (let start = 0; start < total; start += CHUNK_BYTES) {
    const end = Math.min(start + CHUNK_BYTES, total) - 1;
    // The upload URL is pre-authenticated; sending the bearer token to it is wrong.
    const res = await fetch(session.uploadUrl, {
      method: 'PUT',
      headers: { 'Content-Length': String(end - start + 1), 'Content-Range': `bytes ${start}-${end}/${total}` },
      body: file.bytes.subarray(start, end + 1)
    });
    if (!res.ok) throw new Error(`upload of ${file.name} failed at byte ${start}: HTTP ${res.status}`);
    if (res.status === 200 || res.status === 201) item = await res.json();
  }
  if (!item) throw new Error(`upload of ${file.name} did not return the stored file`);
  return item;
}

/**
 * Upload validated files into the channel's SharePoint folder.
 * @returns {Promise<Array<{id, contentType, contentUrl, name}>>} reference attachments
 */
async function uploadToChannel(accessToken, teamId, channelId, files) {
  if (!files.length) return [];
  const folder = await callGraphAPI(accessToken, 'GET', `teams/${teamId}/channels/${channelId}/filesFolder`);
  const driveId = folder && folder.parentReference && folder.parentReference.driveId;
  if (!driveId || !folder.id) throw new Error('the channel has no SharePoint files folder');
  const attachments = [];
  for (const file of files) {
    const item = file.bytes.length <= SIMPLE_UPLOAD_MAX
      ? await callGraphAPI(
        accessToken, 'PUT',
        `drives/${driveId}/items/${folder.id}:/${encodeURIComponent(file.name)}:/content`,
        file.bytes,
        { '@microsoft.graph.conflictBehavior': 'rename' },
        { 'Content-Type': 'application/octet-stream' })
      : await _uploadSession(accessToken, driveId, folder.id, file);
    const guid = /\{?([0-9a-fA-F-]{36})\}?/.exec(String(item.eTag || ''));
    if (!guid || !item.webUrl) throw new Error(`uploaded ${file.name} but Graph returned no eTag/webUrl`);
    attachments.push({ id: guid[1].toLowerCase(), contentType: 'reference', contentUrl: item.webUrl, name: item.name });
  }
  return attachments;
}

/** Place `<attachment>` tags at {{file:N}} markers; the rest go at the end. */
function placeFileTags(content, attachments) {
  let body = content;
  const trailing = [];
  attachments.forEach((a, i) => {
    const tag = `<attachment id="${a.id}"></attachment>`;
    const marker = `{{file:${i + 1}}}`;
    if (body.includes(marker)) body = body.split(marker).join(tag);
    else trailing.push(tag);
  });
  return body + trailing.join('');
}

module.exports = {
  validateFiles, uploadToChannel, placeFileTags,
  MAX_FILES, MAX_FILE_BYTES, SIMPLE_UPLOAD_MAX, SUPPORTED
};
