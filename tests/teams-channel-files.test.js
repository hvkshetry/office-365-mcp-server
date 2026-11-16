const { describe, it, expect, beforeEach, afterAll } = require('@jest/globals');
const fs = require('fs');
const os = require('os');
const path = require('path');

jest.mock('../auth', () => ({ ensureAuthenticated: jest.fn().mockResolvedValue('token') }));
jest.mock('../utils/graph-api');
jest.mock('../policy', () => ({ enforceTeamsPolicy: jest.fn().mockResolvedValue(null) }));

const { callGraphAPI } = require('../utils/graph-api');
const { enforceTeamsPolicy } = require('../policy');
const { validateFiles, SIMPLE_UPLOAD_MAX } = require('../teams/consolidated/channel_files');
const teamsTools = require('../teams');

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-files-'));
const write = (name, bytes) => { const p = path.join(dir, name); fs.writeFileSync(p, bytes); return p; };
// Minimal stand-ins: the check reads the zip magic and the uncompressed part names.
const ooxml = (...parts) => Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]),
  Buffer.from(['[Content_Types].xml', ...parts].join('\0'))]);

const pdf = write('report.pdf', Buffer.from('%PDF-1.7\n1 0 obj\n'));
const xlsx = write('lab.xlsx', ooxml('xl/workbook.xml', 'xl/worksheets/sheet1.xml'));
const docx = write('memo.docx', ooxml('word/document.xml'));
const pptx = write('deck.pptx', ooxml('ppt/presentation.xml'));
const xlsm = write('macro.xlsx', ooxml('xl/workbook.xml', 'xl/vbaProject.bin'));
const envAsPdf = write('secrets.pdf', 'API_KEY=not-a-document\n');
const docxAsXlsx = write('wrong.xlsx', ooxml('word/document.xml'));
const csv = write('data.csv', 'a,b\n1,2\n');

const ITEM = name => ({ eTag: '"{6A4E1C2B-1111-2222-3333-444455556666},1"',
  webUrl: `https://t.sharepoint.com/sites/ops/Shared%20Documents/Minster/${name}`, name });

afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

describe('validateFiles', () => {
  it('accepts PDF, Word, Excel and PowerPoint by content', () => {
    const r = validateFiles([{ path: pdf }, { path: xlsx }, { path: docx }, { path: pptx }]);
    expect(r.error).toBeUndefined();
    expect(r.files.map(f => f.name)).toEqual(['report.pdf', 'lab.xlsx', 'memo.docx', 'deck.pptx']);
  });

  it.each([
    [envAsPdf, /not a PDF/],
    [docxAsXlsx, /not a \.xlsx/],
    [xlsm, /macro-enabled/],
    [csv, /unsupported type \.csv/]
  ])('refuses %s', (p, why) => {
    expect(validateFiles([{ path: p }]).error).toMatch(why);
  });

  it('refuses relative paths, renamed extensions and too many files', () => {
    expect(validateFiles([{ path: 'report.pdf' }]).error).toMatch(/absolute path/);
    expect(validateFiles([{ path: pdf, name: 'report.xlsx' }]).error).toMatch(/must keep/);
    expect(validateFiles(Array(6).fill({ path: pdf })).error).toMatch(/At most 5/);
  });
});

describe('teams_channel files', () => {
  const tool = teamsTools.find(t => t.name === 'teams_channel');

  beforeEach(() => {
    jest.clearAllMocks();
    callGraphAPI.mockImplementation(async (token, method, url) => {
      if (url.endsWith('/filesFolder')) return { id: 'F1', parentReference: { driveId: 'D1' } };
      if (method === 'PUT') return ITEM(decodeURIComponent(url.split(':/')[1].split(':')[0]));
      return { id: 'm-1' };
    });
  });

  it('uploads to the channel folder, then posts a reference attachment at its marker', async () => {
    const res = await tool.handler({ operation: 'reply_to_message', teamId: 't', channelId: 'c', messageId: 'm',
      content: '<p>Here is the workbook:</p>{{file:1}}', files: [{ path: xlsx }] });
    expect(enforceTeamsPolicy).toHaveBeenCalled();
    const calls = callGraphAPI.mock.calls;
    expect(calls[0].slice(1, 3)).toEqual(['GET', 'teams/t/channels/c/filesFolder']);
    expect(calls[1][1]).toBe('PUT');
    expect(calls[1][2]).toBe('drives/D1/items/F1:/lab.xlsx:/content');
    expect(calls[1][4]).toEqual({ '@microsoft.graph.conflictBehavior': 'rename' });
    const [, method, url, payload] = calls[2];
    expect([method, url]).toEqual(['POST', 'teams/t/channels/c/messages/m/replies']);
    const id = '6a4e1c2b-1111-2222-3333-444455556666';
    expect(payload.body).toEqual({ contentType: 'html',
      content: `<p>Here is the workbook:</p><attachment id="${id}"></attachment>` });
    expect(payload.attachments).toEqual([{ id, contentType: 'reference', name: 'lab.xlsx',
      contentUrl: 'https://t.sharepoint.com/sites/ops/Shared%20Documents/Minster/lab.xlsx' }]);
    expect(payload._uploaded).toBeUndefined();
    expect(res.content[0].text).toMatch(/Files attached: lab\.xlsx/);
  });

  it('files and images travel together', async () => {
    const png = write('t.png', Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.from('x')]));
    await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'Report {{image:1}}', images: [{ path: png }], files: [{ path: pdf }] });
    const payload = callGraphAPI.mock.calls.at(-1)[3];
    expect(payload.hostedContents).toHaveLength(1);
    expect(payload.attachments).toHaveLength(1);
    expect(payload.body.content).toMatch(/hostedContents\/1.*<attachment id=/);
  });

  it('a bad file anywhere in the list means nothing is uploaded or posted', async () => {
    const res = await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'x', files: [{ path: pdf }, { path: envAsPdf }] });
    expect(callGraphAPI).not.toHaveBeenCalled();
    expect(res.content[0].text).toMatch(/files\[1\]: not a PDF/);
  });

  it('an upload failure posts nothing', async () => {
    callGraphAPI.mockImplementation(async (token, method, url) => {
      if (url.endsWith('/filesFolder')) return { id: 'F1', parentReference: { driveId: 'D1' } };
      throw new Error('403 accessDenied');
    });
    const res = await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'x', files: [{ path: docx }] });
    expect(callGraphAPI.mock.calls.some(c => c[1] === 'POST')).toBe(false);
    expect(res.content[0].text).toMatch(/nothing was posted: 403/);
  });

  it('a large file goes through an upload session without the bearer token', async () => {
    const big = write('big.pdf', Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.alloc(SIMPLE_UPLOAD_MAX + 10)]));
    callGraphAPI.mockImplementation(async (token, method, url) => {
      if (url.endsWith('/filesFolder')) return { id: 'F1', parentReference: { driveId: 'D1' } };
      if (url.endsWith(':/createUploadSession')) return { uploadUrl: 'https://upload.example/s1' };
      return { id: 'm-2' };
    });
    const seen = [];
    global.fetch = jest.fn(async (u, init) => {
      seen.push(init.headers);
      const [, , end, total] = /bytes (\d+)-(\d+)\/(\d+)/.exec(init.headers['Content-Range']);
      const last = Number(end) + 1 === Number(total);
      return { ok: true, status: last ? 201 : 202, json: async () => ITEM('big.pdf') };
    });
    const res = await tool.handler({ operation: 'create_message', teamId: 't', channelId: 'c',
      content: 'x', files: [{ path: big }] });
    expect(seen.length).toBeGreaterThan(1);
    expect(seen.every(h => !('Authorization' in h))).toBe(true);
    expect(res.content[0].text).toMatch(/Files attached: big\.pdf/);
    delete global.fetch;
  });
});
