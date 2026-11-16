/**
 * Inline images for Teams channel posts and replies.
 *
 * Graph carries an image in a channel message as a hostedContent: the bytes go in
 * `hostedContents` under a temporary id, and the HTML body points at it with
 * <img src="../hostedContents/{id}/$value">. Nothing is uploaded to SharePoint and
 * no file link is shared, so the image is readable by exactly the channel's
 * audience, the same audience as the text.
 *
 * Callers pass local file paths, not base64, so a chart does not cost the model
 * tens of thousands of tokens. Because this reads local files on the caller's
 * behalf, a file is accepted only if its leading bytes identify it as PNG, JPEG
 * or GIF: a key, an .env or a document cannot be smuggled out as an "image".
 */

const fs = require('fs');
const path = require('path');

const MAX_IMAGES = 4;
const MAX_IMAGE_BYTES = 3 * 1024 * 1024;

function _imageType(buf) {
  if (buf.length >= 8 && buf.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    return 'image/png';
  }
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'image/jpeg';
  if (buf.length >= 6 && ['GIF87a', 'GIF89a'].includes(buf.subarray(0, 6).toString('latin1'))) return 'image/gif';
  return null;
}

function _escapeAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Validate `images` and fold them into a message payload.
 *
 * @param {string} content - HTML body. `{{image:N}}` (1-based) places image N there;
 *   images without a placeholder are appended at the end, in order.
 * @param {Array<{path: string, alt?: string}>} images
 * @returns {{content: string, hostedContents: Array}|{error: string}}
 */
function buildHostedImages(content, images) {
  if (!Array.isArray(images) || images.length === 0) return { content, hostedContents: [] };
  if (images.length > MAX_IMAGES) return { error: `At most ${MAX_IMAGES} images per message.` };

  const hostedContents = [];
  const tags = [];
  for (const [i, image] of images.entries()) {
    const p = image && image.path;
    if (typeof p !== 'string' || !path.isAbsolute(p)) {
      return { error: `images[${i}].path must be an absolute path to a PNG, JPEG or GIF file.` };
    }
    let stat;
    try {
      stat = fs.statSync(p);
    } catch {
      return { error: `images[${i}]: file not found: ${p}` };
    }
    if (!stat.isFile()) return { error: `images[${i}]: not a regular file: ${p}` };
    if (stat.size > MAX_IMAGE_BYTES) {
      return { error: `images[${i}]: ${stat.size} bytes exceeds the ${MAX_IMAGE_BYTES}-byte limit.` };
    }
    const bytes = fs.readFileSync(p);
    const contentType = _imageType(bytes);
    if (!contentType) return { error: `images[${i}]: not a PNG, JPEG or GIF image: ${p}` };
    const id = String(i + 1);
    hostedContents.push({ '@microsoft.graph.temporaryId': id, contentBytes: bytes.toString('base64'), contentType });
    const alt = _escapeAttr(image.alt || path.basename(p));
    tags.push(`<img src="../hostedContents/${id}/$value" alt="${alt}" style="max-width:100%">`);
  }

  let body = content;
  const trailing = [];
  tags.forEach((tag, i) => {
    const marker = `{{image:${i + 1}}}`;
    if (body.includes(marker)) body = body.split(marker).join(tag);
    else trailing.push(`<p>${tag}</p>`);
  });
  return { content: body + trailing.join(''), hostedContents };
}

module.exports = { buildHostedImages, MAX_IMAGES, MAX_IMAGE_BYTES };
