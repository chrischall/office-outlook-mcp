import { describe, expect, it, vi } from 'vitest';
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

async function call(client: OutlookClient, args: Record<string, unknown> = { messageId: 'm 1', attachmentId: 'a/1' }) {
  const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
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
