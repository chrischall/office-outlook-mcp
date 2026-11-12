import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_get_attachment` with `convert`: a PDF, workbook or deck comes back
 * as content the MODEL can read (Markdown, CSV, images) inside the untrusted
 * envelope, while the original is still handed over as a link or saved file.
 * Every fixture is built in-test — no binary files.
 */

const FILE = '#Microsoft.OutlookServices.FileAttachment';
const b64 = (b: Uint8Array | string) => Buffer.from(b).toString('base64');
const enc = new TextEncoder();

/** A minimal valid PDF, one page per entry (an empty entry is a page with no text layer). */
function makePdf(pages: string[]): Buffer {
  const esc = (s: string) => s.replace(/[\\()]/g, (c) => `\\${c}`);
  const objects: string[] = [];
  const pageObj = (i: number) => 4 + i * 2;
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[2] = `<< /Type /Pages /Kids [${pages.map((_, i) => `${pageObj(i)} 0 R`).join(' ')}] /Count ${pages.length} >>`;
  objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';
  pages.forEach((t, i) => {
    const stream = t === '' ? '' : `BT\n/F1 12 Tf\n72 720 Td\n(${esc(t)}) Tj\nET`;
    objects[pageObj(i)] =
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${pageObj(i) + 1} 0 R >>`;
    objects[pageObj(i) + 1] = `<< /Length ${Buffer.byteLength(stream, 'latin1')} >>\nstream\n${stream}\nendstream`;
  });
  let out = '%PDF-1.4\n';
  const offsets: number[] = [];
  for (let id = 1; id < objects.length; id++) {
    offsets[id] = Buffer.byteLength(out, 'latin1');
    out += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) out += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/** A workbook with two sheets. */
async function makeXlsx(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('Budget');
  a.addRow(['item', 'cost']);
  a.addRow(['rent', 1200]);
  wb.addWorksheet('Notes').addRow(['hello, world']);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** A 1×1 PNG. */
const PNG = Buffer.from(
  '89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c63000100000500010d0a2db40000000049454e44ae426082',
  'hex',
);

/** A one-slide deck with a title, a bullet and one embedded picture. */
function makePptx(): Buffer {
  const NS =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (list: { id: string; type: string; target: string }[]) =>
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list
      .map((r) => `<Relationship Id="${r.id}" Type="${REL}/${r.type}" Target="${r.target}"/>`)
      .join('')}</Relationships>`;
  const sp = (ph: string, text: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const pic =
    '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill><p:spPr/></p:pic>';
  const slide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${sp('<p:ph type="title"/>', 'Quarterly review')}${sp('<p:ph idx="1"/>', 'Revenue up')}${pic}</p:spTree></p:cSld></p:sld>`;
  return Buffer.from(
    zipSync({
      'ppt/presentation.xml': enc.encode(
        `<?xml version="1.0" encoding="UTF-8"?><p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rIdS1"/></p:sldIdLst></p:presentation>`,
      ),
      'ppt/_rels/presentation.xml.rels': enc.encode(rels([{ id: 'rIdS1', type: 'slide', target: 'slides/slide1.xml' }])),
      'ppt/slides/slide1.xml': enc.encode(slide),
      'ppt/slides/_rels/slide1.xml.rels': enc.encode(rels([{ id: 'rIdImg', type: 'image', target: '../media/image1.png' }])),
      'ppt/media/image1.png': new Uint8Array(PNG),
    }),
  );
}

type Att = Record<string, unknown>;

/** Metadata (no bytes) when the request $selects; the whole record otherwise. */
function stub(att: Att) {
  const get = vi.fn(async (_p: string, o?: { query?: Record<string, unknown> }) => {
    if (o?.query?.$select !== undefined) {
      const { ContentBytes: _drop, ...meta } = att;
      return meta;
    }
    return att;
  });
  const client = { get, write: vi.fn() } as unknown as OutlookClient;
  return { client, get };
}

const fileAtt = (name: string, type: string, bytes: Buffer, extra: Att = {}) =>
  stub({ '@odata.type': FILE, Id: 'a/1', Name: name, ContentType: type, Size: bytes.length, ContentBytes: b64(bytes), ...extra });

let home: string;
beforeEach(() => {
  home = mkdtempSync(join(tmpdir(), 'outlook-conv-home-'));
});
afterEach(() => {
  vi.unstubAllGlobals();
  rmSync(home, { recursive: true, force: true });
});
const localEnv = (extra: Record<string, string> = {}) => ({ HOME: home, ...extra });
const BLOB = { MCP_BLOB_BASE_URL: 'https://host.example/b/reg_0123456789abcdef01234567', MCP_BLOB_SIGNING_KEY: 'a2V5' };

function fakeStore() {
  const puts: { url: string; bytes: Buffer }[] = [];
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string, init?: RequestInit) => {
      puts.push({ url, bytes: Buffer.from(await new Response(init!.body).arrayBuffer()) });
      return new Response(null, { status: 201 });
    }),
  );
  return puts;
}

type Block = { type: string; text?: string; data?: string; mimeType?: string };
type Caption = {
  untrusted_content?: boolean;
  name?: string;
  inlined?: boolean;
  via?: string;
  path?: string;
  downloadUrl?: string;
  hint?: string;
  conversion_error?: string;
  converted?: { format: string; count?: number; parts: string[]; notes: string[] };
};

async function call(client: OutlookClient, extra: Record<string, unknown> = {}, env: Record<string, string> = localEnv()) {
  const h = await createTestHarness((s: McpServer) => registerMailTools(s, client, { env }));
  const res = await h.callTool('outlook_get_attachment', { messageId: 'm1', attachmentId: 'a/1', ...extra });
  await h.close();
  const content = res.content as Block[];
  const last = content[content.length - 1]!;
  const caption = JSON.parse(last.text!) as Caption;
  return { res, content, caption };
}

const texts = (c: Block[]) => c.filter((b) => b.type === 'text').map((b) => b.text!);

describe('outlook_get_attachment — convert', () => {
  it('reads a PDF as Markdown inside the untrusted envelope, and still saves the original locally', async () => {
    const pdf = makePdf(['Hello invoice total 42']);
    const dir = join(home, 'out');
    const { res, content, caption } = await call(fileAtt('report.pdf', 'application/pdf', pdf).client, {}, localEnv({ OUTLOOK_OUTPUT_DIR: dir }));
    expect(res.isError).toBeFalsy();
    // The envelope comes first, before any third-party text.
    const head = JSON.parse(content[0]!.text!) as { untrusted_content?: boolean; note?: string };
    expect(head.untrusted_content).toBe(true);
    expect(head.note).toMatch(/written by other people/);
    // Then the converted Markdown.
    const md = texts(content).find((t) => t.startsWith('--- report.md (text/markdown) ---'));
    expect(md).toContain('## Page 1');
    expect(md).toContain('Hello invoice total 42');
    // The caption closes the result, enveloped, and describes the conversion.
    expect(caption).toMatchObject({ untrusted_content: true, name: 'report.pdf', via: 'disk', inlined: false });
    // One short line is a sparse text layer, so auto also renders the page.
    expect(caption.converted).toMatchObject({ format: 'pdf', count: 1 });
    expect(caption.converted?.parts[0]).toBe('report.md');
    expect(content.some((b) => b.type === 'image')).toBe(true);
    expect(Array.isArray(caption.converted?.notes)).toBe(true);
    // The original is saved, never inlined as a resource.
    expect(readFileSync(caption.path!)).toEqual(pdf);
    expect(content.some((b) => b.type === 'resource')).toBe(false);
  });

  it('on a hosted server, links the original and returns the Markdown — never an embedded resource', async () => {
    const puts = fakeStore();
    const pdf = makePdf(['Quarterly numbers']);
    const { content, caption } = await call(fileAtt('q.pdf', 'application/pdf', pdf).client, {}, { ...BLOB, HOME: home });
    expect(caption.via).toBe('link');
    expect(caption.downloadUrl).toMatch(/^https:\/\/host\.example\//);
    expect(puts).toHaveLength(1);
    expect(puts[0]!.bytes.equals(pdf)).toBe(true);
    expect(texts(content).some((t) => t.includes('Quarterly numbers'))).toBe(true);
    expect(content.some((b) => b.type === 'resource')).toBe(false);
  });

  it('converts even with nowhere to hand the original over, and says the original is not attached', async () => {
    const s = fileAtt('report.pdf', 'application/pdf', makePdf(['No route here']));
    const { content, caption } = await call(s.client);
    expect(texts(content).some((t) => t.includes('No route here'))).toBe(true);
    expect(caption.converted?.format).toBe('pdf');
    expect(caption).not.toHaveProperty('via');
    expect(caption.hint).toMatch(/OUTLOOK_OUTPUT_DIR/);
    expect(s.get).toHaveBeenCalledTimes(2);
  });

  it('hosted without a blob store: converts, says why there is no link, and writes no disk', async () => {
    const s = fileAtt('report.pdf', 'application/pdf', makePdf(['Hosted no store']));
    const { content, caption } = await call(s.client, {}, { MCP_DATA_DIR: join(home, 'data'), HOME: home });
    expect(texts(content).some((t) => t.includes('Hosted no store'))).toBe(true);
    expect(caption).not.toHaveProperty('via');
    expect(caption.hint).toMatch(/blob store/i);
  });

  it('converts a file over the 5 MB inline cap (the cap is for inlining, not conversion input)', async () => {
    const pdf = makePdf(['Big one']);
    const s = fileAtt('big.pdf', 'application/pdf', pdf, { Size: 6 * 1024 * 1024 });
    const { content, caption } = await call(s.client);
    expect(texts(content).some((t) => t.includes('Big one'))).toBe(true);
    expect(caption.converted?.format).toBe('pdf');
  });

  it('still refuses to download a convertible file over 50 MB', async () => {
    const s = fileAtt('huge.pdf', 'application/pdf', makePdf(['x']), { Size: 50 * 1024 * 1024 + 1 });
    const { caption } = await call(s.client);
    expect(caption.hint).toMatch(/50 MB/);
    expect(caption).not.toHaveProperty('converted');
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it('reads an Excel workbook as one CSV per sheet', async () => {
    const xlsx = await makeXlsx();
    const { content, caption } = await call(
      fileAtt('budget.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsx).client,
      {},
      localEnv({ OUTLOOK_OUTPUT_DIR: join(home, 'out') }),
    );
    const csv = texts(content);
    expect(csv.find((t) => t.startsWith('--- Budget.csv (text/csv) ---'))).toMatch(/item,cost\r?\nrent,1200/);
    expect(csv.find((t) => t.startsWith('--- Notes.csv (text/csv) ---'))).toContain('"hello, world"');
    expect(caption.converted).toMatchObject({ format: 'xlsx', parts: ['Budget.csv', 'Notes.csv'] });
    expect(caption.via).toBe('disk');
  });

  it('reads a PowerPoint deck as slide Markdown plus its pictures (extension wins for octet-stream)', async () => {
    const { content, caption } = await call(
      fileAtt('deck.pptx', 'application/octet-stream', makePptx()).client,
      {},
      localEnv({ OUTLOOK_OUTPUT_DIR: join(home, 'out') }),
    );
    const md = texts(content).find((t) => t.startsWith('--- deck.md (text/markdown) ---'));
    expect(md).toContain('## Slide 1: Quarterly review');
    expect(md).toContain('Revenue up');
    expect(content.find((b) => b.type === 'image')).toMatchObject({ type: 'image', mimeType: 'image/png', data: b64(PNG) });
    expect(caption.converted?.format).toBe('pptx');
  });

  it('passes the mode through: images renders PDF pages, text never does', async () => {
    const pdf = makePdf(['Page with words']);
    const imgs = await call(fileAtt('r.pdf', 'application/pdf', pdf).client, { convert: 'images' });
    expect(imgs.content.some((b) => b.type === 'image' && /^image\/(png|jpeg)$/.test(b.mimeType ?? ''))).toBe(true);
    expect(imgs.caption.converted?.parts).toEqual([expect.stringMatching(/^r-page-1\.(png|jpg)$/)]);
    // A page with no text layer is a scan: auto would render it, text must not.
    const text = await call(fileAtt('r.pdf', 'application/pdf', makePdf([''])).client, { convert: 'text' });
    expect(text.content.some((b) => b.type === 'image')).toBe(false);
    expect(text.caption.converted?.format).toBe('pdf');
  });

  it('convert: "none" keeps the old behaviour — no conversion, bytes not fetched without a route', async () => {
    const s = fileAtt('report.pdf', 'application/pdf', makePdf(['skip me']));
    const { content, caption } = await call(s.client, { convert: 'none' });
    expect(caption).not.toHaveProperty('converted');
    expect(caption.hint).toMatch(/not inlined/i);
    expect(texts(content).some((t) => t.includes('skip me'))).toBe(false);
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it('a corrupt file does not fail the tool: it falls back and reports conversion_error', async () => {
    const bad = Buffer.from('%PDF-1.7\nthis is not really a pdf\n%%EOF\n');
    const dir = join(home, 'out');
    const { res, content, caption } = await call(fileAtt('bad.pdf', 'application/pdf', bad).client, {}, localEnv({ OUTLOOK_OUTPUT_DIR: dir }));
    expect(res.isError).toBeFalsy();
    expect(caption.conversion_error).toEqual(expect.any(String));
    expect(caption).not.toHaveProperty('converted');
    expect(caption.via).toBe('disk');
    expect(readdirSync(dir)).toHaveLength(1);
    // Fallback is the plain metadata answer: one enveloped block.
    expect(content).toHaveLength(1);
  });

  it('bytes that are not the declared format fall back with conversion_error too', async () => {
    const { caption } = await call(fileAtt('fake.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', Buffer.from('plain text')).client);
    expect(caption.conversion_error).toMatch(/xlsx|Excel|not/i);
    expect(caption.hint).toMatch(/OUTLOOK_OUTPUT_DIR/);
  });

  it('leaves non-convertible types alone', async () => {
    const s = fileAtt('a.zip', 'application/zip', Buffer.from('PK\x03\x04zip'));
    const { caption } = await call(s.client);
    expect(caption).not.toHaveProperty('converted');
    expect(caption).not.toHaveProperty('conversion_error');
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it('documents convert in the schema and the description', async () => {
    const h = await createTestHarness((srv: McpServer) => registerMailTools(srv, stub({}).client));
    const { tools } = await h.client.listTools();
    await h.close();
    const tool = tools.find((t) => t.name === 'outlook_get_attachment');
    const prop = (tool?.inputSchema as { properties?: Record<string, { enum?: string[]; description?: string }> }).properties?.convert;
    expect(prop?.enum).toEqual(['auto', 'text', 'images', 'none']);
    expect(prop?.description).toMatch(/CSV/);
    expect(tool?.description).toMatch(/PDF/);
    expect(tool?.description).toMatch(/convert/);
  });
});
