import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_get_attachment` — one attachment's content, in the form a model can
 * use: an image as MCP image content, text-like files decoded (capped), any
 * other binary as metadata only. Fixtures are synthetic placeholder bytes.
 */

const FILE = '#Microsoft.OutlookServices.FileAttachment';
const b64 = (s: string | Buffer) => Buffer.from(s).toString('base64');

type Att = Record<string, unknown>;

/**
 * The tool reads metadata first (no bytes), then the full attachment only when
 * it will inline it. The stub answers both from one record, dropping
 * ContentBytes when the request $selects metadata.
 */
function stub(att: Att) {
  const get = vi.fn(async (_p: string, o?: { query?: Record<string, unknown> }) => {
    if (o?.query?.$select !== undefined) {
      const { ContentBytes: _drop, ...meta } = att;
      return meta;
    }
    return att;
  });
  const write = vi.fn(async () => ({}));
  const client = { get, write } as unknown as OutlookClient;
  return { client, get, write };
}

/**
 * A local install with NO download directory: a HOME without ~/Downloads and
 * no OUTLOOK_OUTPUT_DIR. The default for every test that does not name one,
 * so nothing here can write into the developer's real ~/Downloads.
 */
let home: string;
beforeEach(() => {
  home = mkdtempSync(join(tmpdir(), 'outlook-att-home-'));
});
afterEach(() => {
  vi.unstubAllGlobals();
  rmSync(home, { recursive: true, force: true });
});
const localEnv = (extra: Record<string, string> = {}) => ({ HOME: home, ...extra });

async function call(
  client: OutlookClient,
  args: Record<string, unknown> = { messageId: 'm 1', attachmentId: 'a/1' },
  env: Record<string, string | undefined> = localEnv(),
) {
  const h = await createTestHarness((s: McpServer) => registerMailTools(s, client, { env }));
  const res = await h.callTool('outlook_get_attachment', args);
  await h.close();
  return res;
}

type Out = {
  untrusted_content?: boolean;
  name?: string;
  contentType?: string;
  size?: number;
  kind?: string;
  inlined?: boolean;
  text?: string;
  textTruncated?: boolean;
  hint?: string;
  via?: string;
  downloadUrl?: string;
  downloadExpiresAt?: string;
  path?: string;
  deliveryNote?: string;
};

describe('outlook_get_attachment', () => {
  it('returns an image as MCP image content, beside its metadata', async () => {
    const bytes = b64('placeholder-png-bytes');
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'chart.png', ContentType: 'image/png', Size: 21, ContentBytes: bytes });
    const res = await call(s.client);
    const content = res.content as { type: string; data?: string; mimeType?: string }[];
    const image = content.find((c) => c.type === 'image');
    expect(image).toEqual({ type: 'image', data: bytes, mimeType: 'image/png' });
    const meta = parseToolResult<Out>(res);
    expect(meta).toMatchObject({ untrusted_content: true, name: 'chart.png', kind: 'file', inlined: true });
    // Both ids are path-encoded; the metadata probe selects no bytes.
    expect(s.get.mock.calls[0][0]).toBe('/me/messages/m%201/attachments/a%2F1');
    expect(String(s.get.mock.calls[0][1]?.query?.$select)).not.toContain('ContentBytes');
  });

  it.each([
    ['text/plain', 'notes.txt'],
    ['text/csv', 'rows.csv'],
    ['application/json', 'data.json'],
    ['application/xml', 'feed.xml'],
    ['text/calendar', 'invite.ics'],
    ['application/octet-stream', 'invite.ics'],
  ])('decodes %s (%s) to text inside the untrusted envelope', async (type, name) => {
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: name, ContentType: type, Size: 11, ContentBytes: b64('hello world') });
    const res = await call(s.client);
    expect((res.content as { type: string }[]).every((c) => c.type === 'text')).toBe(true);
    const out = parseToolResult<Out>(res);
    expect(out.untrusted_content).toBe(true);
    expect(out.text).toBe('hello world');
    expect(out.inlined).toBe(true);
    expect(out).not.toHaveProperty('textTruncated');
  });

  it('caps decoded text at 200 KB and says so', async () => {
    const big = 'x'.repeat(300 * 1024);
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'big.txt', ContentType: 'text/plain', Size: big.length, ContentBytes: b64(big) });
    const out = parseToolResult<Out>(await call(s.client));
    expect(out.text).toHaveLength(200 * 1024);
    expect(out.textTruncated).toBe(true);
  });

  it('never splits a multi-byte UTF-8 character at the 200 KB cut', async () => {
    // 'é' is two bytes; placed so the cut lands after its first byte.
    const big = `${'x'.repeat(200 * 1024 - 1)}é${'y'.repeat(100)}`;
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'big.txt', ContentType: 'text/plain', Size: 1, ContentBytes: b64(big) });
    const out = parseToolResult<Out>(await call(s.client));
    expect(out.text).toBe('x'.repeat(200 * 1024 - 1));
    expect(out.text).not.toContain('\uFFFD');
    expect(out.textTruncated).toBe(true);
  });

  it('refuses to inline content over 5 MB even when the metadata size understates it', async () => {
    const huge = b64(Buffer.alloc(5 * 1024 * 1024 + 3, 0x61));
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'huge.txt', ContentType: 'text/plain', Size: 10, ContentBytes: huge });
    const from = vi.spyOn(Buffer, 'from');
    try {
      const out = parseToolResult<Out>(await call(s.client));
      expect(out.inlined).toBe(false);
      expect(out.hint).toMatch(/5 MB/);
      expect(out).not.toHaveProperty('text');
      // Judged from the base64 length, before decoding it.
      expect(from.mock.calls.some((c) => (c as unknown[])[0] === huge && (c as unknown[])[1] === 'base64')).toBe(false);
    } finally {
      from.mockRestore();
    }
  });

  it('returns an SVG as decoded text, not image content', async () => {
    const svg = '<svg xmlns="http://www.w3.org/2000/svg"><text>hi</text></svg>';
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'logo.svg', ContentType: 'image/svg+xml', Size: svg.length, ContentBytes: b64(svg) });
    const res = await call(s.client);
    expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(false);
    const out = parseToolResult<Out>(res);
    expect(out).toMatchObject({ inlined: true, text: svg, untrusted_content: true });
  });

  it('returns other binary types as metadata only, never fetching the bytes', async () => {
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'report.pdf', ContentType: 'application/pdf', Size: 1234, ContentBytes: b64('pdf') });
    const out = parseToolResult<Out>(await call(s.client));
    expect(out).toMatchObject({ name: 'report.pdf', contentType: 'application/pdf', size: 1234, inlined: false });
    expect(out).not.toHaveProperty('contentBytes');
    expect(out).not.toHaveProperty('text');
    expect(out.hint).toMatch(/not inlined/i);
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it('never inlines an image or text file above 5 MB', async () => {
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'huge.png', ContentType: 'image/png', Size: 6 * 1024 * 1024 });
    const res = await call(s.client);
    expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(false);
    const out = parseToolResult<Out>(res);
    expect(out.inlined).toBe(false);
    expect(out.hint).toMatch(/5 MB/);
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it.each([
    ['#Microsoft.OutlookServices.ItemAttachment', 'item', /attached Outlook item/i],
    ['#Microsoft.OutlookServices.ReferenceAttachment', 'reference', /link/i],
  ])('explains why a %s has no bytes', async (type, kind, why) => {
    const s = stub({ '@odata.type': type, Id: 'a/1', Name: 'Fwd thing', ContentType: null, Size: 99 });
    const out = parseToolResult<Out>(await call(s.client));
    expect(out).toMatchObject({ kind, inlined: false, name: 'Fwd thing' });
    expect(out.hint).toMatch(why);
    expect(s.get).toHaveBeenCalledTimes(1);
  });

  it('reports a file whose bytes did not come back instead of inlining nothing', async () => {
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'n.txt', ContentType: 'text/plain', Size: 3 });
    const out = parseToolResult<Out>(await call(s.client));
    expect(out.inlined).toBe(false);
    expect(out.hint).toMatch(/no content/i);
  });

  it('never writes, and its description says the text is untrusted', async () => {
    const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'n.txt', ContentType: 'text/plain', Size: 1, ContentBytes: b64('n') });
    await call(s.client);
    expect(s.write).not.toHaveBeenCalled();
    const h = await createTestHarness((srv: McpServer) => registerMailTools(srv, s.client));
    const tool = (await h.listTools()).find((t) => t.name === 'outlook_get_attachment');
    await h.close();
    expect(tool?.description).toMatch(/untrusted/i);
    expect(tool?.description).toMatch(/outlook_list_attachments/);
  });
});

/**
 * Handing the BYTES over. A PDF, an Office file or anything over 5 MB used to
 * come back as metadata only, so a caller could never get the file. Now it goes
 * through mcp-utils' `deliverFile`: a signed download link when hosted, a saved
 * file when a local install has a download directory, and — with neither — the
 * old metadata-only answer.
 */
describe('outlook_get_attachment — file delivery', () => {
  const PDF = Buffer.from('%PDF-1.7\nplaceholder\n%%EOF\n');
  const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d]);
  const BLOB = { MCP_BLOB_BASE_URL: 'https://host.example/b/reg_0123456789abcdef01234567', MCP_BLOB_SIGNING_KEY: 'a2V5' };

  /** A fake blob store: the real signer behind a stubbed `fetch` that records every PUT. */
  function fakeStore() {
    const puts: { url: string; type: string; bytes: Buffer }[] = [];
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string, init?: RequestInit) => {
        puts.push({
          url,
          type: new Headers(init!.headers).get('content-type')!,
          bytes: Buffer.from(await new Response(init!.body).arrayBuffer()),
        });
        return new Response(null, { status: 201 });
      }),
    );
    return puts;
  }

  const pdf = (extra: Att = {}) =>
    stub({ '@odata.type': FILE, Id: 'a/1', Name: 'report.pdf', ContentType: 'application/pdf', Size: PDF.length, ContentBytes: b64(PDF), ...extra });

  describe('local install', () => {
    it('saves a PDF under OUTLOOK_OUTPUT_DIR (0600) and keeps the metadata', async () => {
      const dir = join(home, 'out');
      const s = pdf();
      const res = await call(s.client, undefined, localEnv({ OUTLOOK_OUTPUT_DIR: dir }));
      const out = parseToolResult<Out>(res);
      expect(out).toMatchObject({
        untrusted_content: true, name: 'report.pdf', contentType: 'application/pdf', size: PDF.length, kind: 'file', inlined: false, via: 'disk',
      });
      expect(out.path).toBe(join(dir, 'report.pdf'));
      expect(readFileSync(out.path!)).toEqual(PDF);
      expect(statSync(out.path!).mode & 0o777).toBe(0o600);
      expect(out).not.toHaveProperty('downloadUrl');
      expect((res.content as { type: string }[]).every((c) => c.type === 'text')).toBe(true);
    });

    it('never overwrites: a second save gets a fresh name', async () => {
      const dir = join(home, 'out');
      const env = localEnv({ OUTLOOK_OUTPUT_DIR: dir });
      const first = parseToolResult<Out>(await call(pdf().client, undefined, env));
      const second = parseToolResult<Out>(await call(pdf().client, undefined, env));
      expect(second.path).not.toBe(first.path);
      expect(readdirSync(dir)).toHaveLength(2);
    });

    it('defaults to ~/Downloads/office-outlook-mcp when ~/Downloads exists', async () => {
      mkdirSync(join(home, 'Downloads'));
      const out = parseToolResult<Out>(await call(pdf().client));
      expect(out.via).toBe('disk');
      expect(out.path).toBe(join(home, 'Downloads', 'office-outlook-mcp', 'report.pdf'));
    });

    it('with no download directory, answers as before: metadata only, bytes never fetched', async () => {
      const s = pdf();
      const out = parseToolResult<Out>(await call(s.client));
      expect(out).toMatchObject({ inlined: false });
      expect(out).not.toHaveProperty('via');
      expect(out.hint).toMatch(/OUTLOOK_OUTPUT_DIR/);
      expect(s.get).toHaveBeenCalledTimes(1);
    });

    it('saves an image over 5 MB instead of inlining it', { timeout: 30_000 }, async () => {
      const big = Buffer.concat([PNG, Buffer.alloc(5 * 1024 * 1024)]);
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'scan.png', ContentType: 'image/png', Size: big.length, ContentBytes: b64(big) });
      const res = await call(s.client, undefined, localEnv({ OUTLOOK_OUTPUT_DIR: join(home, 'out') }));
      expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(false);
      const out = parseToolResult<Out>(res);
      expect(out).toMatchObject({ via: 'disk', inlined: false });
      expect(readFileSync(out.path!)).toEqual(big);
    });

    it('still inlines a small image as image content, writing nothing', async () => {
      const dir = join(home, 'out');
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'chart.png', ContentType: 'image/png', Size: PNG.length, ContentBytes: b64(PNG) });
      const res = await call(s.client, undefined, localEnv({ OUTLOOK_OUTPUT_DIR: dir }));
      expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(true);
      expect(parseToolResult<Out>(res)).toMatchObject({ inlined: true, via: 'image' });
      expect(() => readdirSync(dir)).toThrow();
    });
  });

  describe('hosted with a blob store', () => {
    it('returns a PDF as a signed download link and never touches disk', async () => {
      const puts = fakeStore();
      const dir = join(home, 'out');
      const s = pdf();
      const res = await call(s.client, undefined, { ...BLOB, MCP_DATA_DIR: join(home, 'data'), HOME: home, OUTLOOK_OUTPUT_DIR: dir });
      const out = parseToolResult<Out>(res);
      expect(out).toMatchObject({ name: 'report.pdf', size: PDF.length, kind: 'file', inlined: false, via: 'link' });
      expect(out.downloadUrl).toMatch(/^https:\/\/host\.example\/b\/reg_0123456789abcdef01234567\/attachments\/.+\/report\.pdf\?exp=\d+&sig=/);
      expect(Date.parse(out.downloadExpiresAt!)).toBeGreaterThan(Date.now());
      expect(out).not.toHaveProperty('path');
      expect(puts).toHaveLength(1);
      expect(puts[0].type).toBe('application/pdf');
      // Exactly the file: a pooled Buffer must not upload the rest of its pool.
      expect(puts[0].bytes.equals(PDF)).toBe(true);
      expect(() => readdirSync(dir)).toThrow();
      expect((res.content as { type: string }[]).every((c) => c.type === 'text')).toBe(true);
    });

    it('inlines a small image as image content without a PUT', async () => {
      const puts = fakeStore();
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'chart.png', ContentType: 'image/png', Size: PNG.length, ContentBytes: b64(PNG) });
      const res = await call(s.client, undefined, { ...BLOB, HOME: home });
      expect((res.content as { type: string }[]).find((c) => c.type === 'image')).toEqual({ type: 'image', data: b64(PNG), mimeType: 'image/png' });
      expect(parseToolResult<Out>(res)).toMatchObject({ inlined: true, via: 'image' });
      expect(puts).toHaveLength(0);
    });

    it('links an image over 5 MB', { timeout: 30_000 }, async () => {
      const puts = fakeStore();
      const big = Buffer.concat([PNG, Buffer.alloc(5 * 1024 * 1024)]);
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'scan.png', ContentType: 'image/png', Size: big.length, ContentBytes: b64(big) });
      const res = await call(s.client, undefined, { ...BLOB, HOME: home });
      expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(false);
      expect(parseToolResult<Out>(res)).toMatchObject({ via: 'link', inlined: false });
      expect(puts).toHaveLength(1);
    });

    it('links a text file too big to decode, even when its metadata size understates it', { timeout: 30_000 }, async () => {
      const puts = fakeStore();
      const huge = Buffer.alloc(5 * 1024 * 1024 + 3, 0x61);
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'huge.txt', ContentType: 'text/plain', Size: 10, ContentBytes: b64(huge) });
      const out = parseToolResult<Out>(await call(s.client, undefined, { ...BLOB, HOME: home }));
      expect(out).toMatchObject({ via: 'link', inlined: false });
      expect(out).not.toHaveProperty('text');
      expect(puts[0].bytes.length).toBe(huge.length);
    });

    it('still decodes a small text file to text, with no PUT', async () => {
      const puts = fakeStore();
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'n.txt', ContentType: 'text/plain', Size: 5, ContentBytes: b64('hello') });
      const out = parseToolResult<Out>(await call(s.client, undefined, { ...BLOB, HOME: home }));
      expect(out).toMatchObject({ text: 'hello', inlined: true });
      expect(out).not.toHaveProperty('via');
      expect(puts).toHaveLength(0);
    });
  });

  describe('hosted without a blob store', () => {
    it('refuses clearly, never fetching the bytes and never writing disk', async () => {
      mkdirSync(join(home, 'Downloads'));
      const s = pdf();
      const out = parseToolResult<Out>(await call(s.client, undefined, { MCP_DATA_DIR: join(home, 'data'), HOME: home }));
      expect(out).toMatchObject({ name: 'report.pdf', inlined: false });
      expect(out).not.toHaveProperty('via');
      expect(out.hint).toMatch(/blob store/i);
      expect(s.get).toHaveBeenCalledTimes(1);
      expect(readdirSync(join(home, 'Downloads'))).toHaveLength(0);
    });

    it('still inlines a small image', async () => {
      const s = stub({ '@odata.type': FILE, Id: 'a/1', Name: 'chart.png', ContentType: 'image/png', Size: PNG.length, ContentBytes: b64(PNG) });
      const res = await call(s.client, undefined, { MCP_DATA_DIR: join(home, 'data'), HOME: home });
      expect((res.content as { type: string }[]).some((c) => c.type === 'image')).toBe(true);
    });
  });

  describe('download cap (50 MB)', () => {
    it('does not download an attachment whose metadata size is over the cap', async () => {
      fakeStore();
      const s = pdf({ Size: 50 * 1024 * 1024 + 1 });
      const out = parseToolResult<Out>(await call(s.client, undefined, { ...BLOB, HOME: home }));
      expect(out).toMatchObject({ inlined: false });
      expect(out).not.toHaveProperty('via');
      expect(out.hint).toMatch(/50 MB/);
      expect(s.get).toHaveBeenCalledTimes(1);
    });

    it('refuses content over the cap when the metadata size understates it, before decoding', { timeout: 30_000 }, async () => {
      const puts = fakeStore();
      const huge = 'A'.repeat(Math.ceil(((50 * 1024 * 1024 + 4) * 4) / 3 / 4) * 4);
      const s = pdf({ Size: 10, ContentBytes: huge });
      const out = parseToolResult<Out>(await call(s.client, undefined, { ...BLOB, HOME: home }));
      expect(out).toMatchObject({ inlined: false });
      expect(out.hint).toMatch(/50 MB/);
      expect(puts).toHaveLength(0);
    });
  });
});
