const { describe, it, expect, beforeEach, afterAll } = require('@jest/globals');
const fs = require('fs');
const os = require('os');
const path = require('path');

jest.mock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('token') }));
jest.mock('../utils/graph-api');
jest.mock('../policy', () => ({ enforceTeamsPolicy: jest.fn().mockResolvedValue(null) }));

const { callGraphAPI } = require('../utils/graph-api');
const { enforceTeamsPolicy } = require('../policy');
const { buildHostedImages, MAX_IMAGE_BYTES } = require('../teams/consolidated/channel_images');
const teamsTools = require('../teams');

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-img-'));
const PNG = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.from('chart')]);
const png = path.join(dir, 'trend.png');
fs.writeFileSync(png, PNG);
const jpg = path.join(dir, 'photo.jpg');
fs.writeFileSync(jpg, Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]));
const secret = path.join(dir, 'fake.png');
fs.writeFileSync(secret, 'API_KEY=not-an-image\n');

afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

describe('buildHostedImages', () => {
  it('places images at their markers and appends the rest', () => {
    const r = buildHostedImages('<p>pH left its band.</p>{{image:1}}<p>Details</p>',
      [{ path: png, alt: 'pH "trend"' }, { path: jpg }]);
    expect(r.error).toBeUndefined();
    expect(r.content).toBe('<p>pH left its band.</p>'
      + '<img src="../hostedContents/1/$value" alt="pH &quot;trend&quot;" style="max-width:100%">'
      + '<p>Details</p><p><img src="../hostedContents/2/$value" alt="photo.jpg" style="max-width:100%"></p>');
    expect(r.hostedContents).toEqual([
      { '@microsoft.graph.temporaryId': '1', contentBytes: PNG.toString('base64'), contentType: 'image/png' },
      expect.objectContaining({ '@microsoft.graph.temporaryId': '2', contentType: 'image/jpeg' })
    ]);
  });

  it('refuses a non-image file whatever its extension', () => {
    expect(buildHostedImages('x', [{ path: secret }]).error).toMatch(/not a PNG, JPEG or GIF/);
  });

  it('refuses relative paths, missing files, directories, oversize files and too many images', () => {
    expect(buildHostedImages('x', [{ path: 'trend.png' }]).error).toMatch(/absolute path/);
    expect(buildHostedImages('x', [{ path: path.join(dir, 'nope.png') }]).error).toMatch(/not found/);
    expect(buildHostedImages('x', [{ path: dir }]).error).toMatch(/not a regular file/);
    const big = path.join(dir, 'big.png');
    fs.writeFileSync(big, Buffer.concat([PNG, Buffer.alloc(MAX_IMAGE_BYTES)]));
    expect(buildHostedImages('x', [{ path: big }]).error).toMatch(/exceeds/);
    expect(buildHostedImages('x', Array(5).fill({ path: png })).error).toMatch(/At most 4/);
  });

  it('passes content through untouched when there are no images', () => {
    expect(buildHostedImages('plain', undefined)).toEqual({ content: 'plain', hostedContents: [] });
  });
});

describe('teams_channel posts with images', () => {
  const tool = teamsTools.find(t => t.name === 'teams_channel');

  beforeEach(() => {
    jest.clearAllMocks();
    callGraphAPI.mockResolvedValue({ id: '123' });
  });

  it('create_message sends hostedContents and an HTML body, after the fence', async () => {
    const res = await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'pH below band {{image:1}}', images: [{ path: png }] });
    expect(enforceTeamsPolicy).toHaveBeenCalled();
    const [, method, url, payload] = callGraphAPI.mock.calls[0];
    expect(method).toBe('POST');
    expect(url).toBe('teams/t/channels/c/messages');
    expect(payload.body.contentType).toBe('html');
    expect(payload.body.content).toContain('../hostedContents/1/$value');
    expect(payload.hostedContents).toHaveLength(1);
    expect(res.content[0].text).toMatch(/ID: 123/);
  });

  it('reply_to_message supports images too', async () => {
    await tool.handler({ operation: 'reply_to_message', teamId: 't', channelId: 'c', messageId: 'm',
      content: '<p>Replay</p>', images: [{ path: png }] });
    const [, , url, payload] = callGraphAPI.mock.calls[0];
    expect(url).toBe('teams/t/channels/c/messages/m/replies');
    expect(payload.hostedContents).toHaveLength(1);
  });

  it('a bad image is refused before anything is posted', async () => {
    const res = await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'x', images: [{ path: secret }] });
    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(res.content[0].text).toMatch(/not a PNG/);
  });

  it('reference attachments keep their id', async () => {
    await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: '<attachment id="abc"></attachment>',
      attachments: [{ id: 'abc', name: 'a.xlsx', contentUrl: 'https://x/a.xlsx', contentType: 'reference' }] });
    expect(callGraphAPI.mock.calls[0][3].attachments[0].id).toBe('abc');
  });
});
