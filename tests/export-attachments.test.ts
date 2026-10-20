import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerExportTools, type ExportToolOptions } from '../src/tools/export.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `attachments: true` on both export tools: file attachments are saved beside
 * the records, item/reference attachments and failures are recorded as
 * skipped, and nothing about one attachment fails the export. Synthetic bytes.
 */

const FILE = '#Microsoft.OutlookServices.FileAttachment';
const ITEM = '#Microsoft.OutlookServices.ItemAttachment';
const REF = '#Microsoft.OutlookServices.ReferenceAttachment';
const b64 = (s: string) => Buffer.from(s).toString('base64');
const MB = 1024 * 1024;

type Att = Record<string, unknown>;

const ATTS: Record<string, Att[]> = {
  'id-1': [
    { '@odata.type': FILE, Id: 'a1', Name: 'report.pdf', ContentType: 'application/pdf', Size: 10, ContentBytes: b64('%PDF-fake') },
    { '@odata.type': FILE, Id: 'a2', Name: 'report.pdf', ContentType: 'application/pdf', Size: 10, ContentBytes: b64('second') },
    { '@odata.type': FILE, Id: 'a3', Name: '../../etc/passwd', ContentType: 'text/plain', Size: 5, ContentBytes: b64('hello') },
    { '@odata.type': ITEM, Id: 'a4', Name: 'Fwd: thing', Size: 100 },
    { '@odata.type': REF, Id: 'a5', Name: 'deck.pptx', Size: 0 },
    { '@odata.type': FILE, Id: 'a6', Name: 'huge.bin', ContentType: 'application/octet-stream', Size: 60 * MB },
    { '@odata.type': FILE, Id: 'a7', Name: 'broken.txt', ContentType: 'text/plain', Size: 4 },
  ],
  'ev-1': [{ '@odata.type': FILE, Id: 'e1', Name: 'agenda.txt', ContentType: 'text/plain', Size: 6, ContentBytes: b64('agenda') }],
};

const msg = (n: number, has: boolean) => ({
  Id: `id-${n}`,
  Subject: `Subject ${n}`,
  ReceivedDateTime: `2026-01-01T00:0${n}:00Z`,
  HasAttachments: has,
  Body: { ContentType: 'Text', Content: 'body' },
});
const MAILBOX = [msg(1, true), msg(2, false)];

const event = {
  Id: 'ev-1',
  Subject: 'Planning',
  Start: { DateTime: '2026-03-01T09:00:00.0000000', TimeZone: 'UTC' },
  End: { DateTime: '2026-03-01T10:00:00.0000000', TimeZone: 'UTC' },
  HasAttachments: true,
  Body: { ContentType: 'Text', Content: 'body' },
};

type Opts = { query?: Record<string, unknown>; text?: boolean; prefer?: string };

function stub() {
  const calls: { path: string; opts?: Opts }[] = [];
  const get = vi.fn(async (path: string, opts?: Opts) => {
    calls.push({ path, opts });
    if (path === '/me') return { EmailAddress: 'me@example.test' };
    if (path === '/me/MailboxSettings') return { TimeZone: 'UTC' };
    if (path.endsWith('/calendarview')) return { value: [event] };
    const one = /^\/me\/(messages|events)\/([^/]+)\/attachments\/([^/]+)$/.exec(path);
    if (one) {
      const id = decodeURIComponent(one[2]);
      const att = ATTS[id]?.find((a) => a.Id === decodeURIComponent(one[3]));
      if (!att || att.Id === 'a7') throw Object.assign(new Error('Outlook 500 on download'), { status: 500 });
      return att;
    }
    const list = /^\/me\/(messages|events)\/([^/]+)\/attachments$/.exec(path);
    if (list) {
      // A listing returns metadata only.
      return { value: (ATTS[decodeURIComponent(list[2])] ?? []).map(({ ContentBytes: _b, ...meta }) => meta) };
    }
    return { value: MAILBOX };
  });
  return { client: { get, getAbsolute: vi.fn(), write: vi.fn() } as unknown as OutlookClient, calls, get };
}

type Result = {
  mode: 'disk' | 'blob';
  has_more?: boolean;
  cursor?: string;
  location: string;
  files: { path?: string; name?: string; bytes: number }[];
  counts: Record<string, number>;
  errors?: string[];
};

let root: string;
beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'outlook-export-att-'));
});
afterEach(() => {
  rmSync(root, { recursive: true, force: true });
});

async function call(tool: string, client: OutlookClient, args: Record<string, unknown>, opts: ExportToolOptions = {}) {
  const h = await createTestHarness((s: McpServer) =>
    registerExportTools(s, client, { env: { OUTLOOK_OUTPUT_DIR: root }, sleep: async () => {}, ...opts }),
  );
  const res = await h.callTool(tool, args);
  await h.close();
  return res;
}

const records = (path: string) =>
  readFileSync(path, 'utf8')
    .split('\n')
    .filter(Boolean)
    .map((l) => JSON.parse(l) as Record<string, unknown>)
    .slice(1);

type Entry = { id?: string; name?: string; kind?: string; path?: string; skipped?: string };

describe('export attachments', () => {
  it('saves file attachments beside the mail records and records the rest as skipped', async () => {
    const s = stub();
    const out = parseToolResult<Result>(await call('outlook_export_messages', s.client, { attachments: true }));
    expect(out.counts).toEqual({ messages: 2, attachments: 3, attachments_skipped: 4 });
    expect(out.errors).toBeUndefined();

    const part = out.files.find((f) => f.path?.endsWith('part-0001.jsonl'))!;
    const [withAtt, without] = records(part.path!);
    expect(without.attachments).toBeUndefined();
    // A message without HasAttachments is never asked for its attachments.
    expect(s.calls.some((c) => c.path.startsWith('/me/messages/id-2/'))).toBe(false);

    const list = withAtt.attachments as Entry[];
    expect(list).toHaveLength(7);
    const saved = list.filter((a) => a.path);
    expect(saved).toHaveLength(3);
    for (const a of saved) expect(a.path).toMatch(/^attachments\/[0-9a-f]{8,}\/[^/]+$/);
    // Same name twice gets a distinct file; a hostile name is reduced to one safe component.
    expect(new Set(saved.map((a) => a.path)).size).toBe(3);
    expect(saved.every((a) => !a.path!.includes('..'))).toBe(true);
    expect(readFileSync(join(out.location, saved[0].path!), 'utf8')).toBe('%PDF-fake');
    expect(readFileSync(join(out.location, saved[1].path!), 'utf8')).toBe('second');

    const skipped = Object.fromEntries(list.filter((a) => a.skipped).map((a) => [a.id, a.skipped!]));
    expect(skipped.a4).toMatch(/item/i);
    expect(skipped.a5).toMatch(/link|reference|cloud/i);
    expect(skipped.a6).toMatch(/too large|50 MB/i);
    expect(skipped.a7).toMatch(/failed/i);
    // The oversize attachment's bytes are never downloaded.
    expect(s.calls.some((c) => c.path.endsWith('/attachments/a6'))).toBe(false);

    // The manifest lists the attachment files too.
    const manifest = JSON.parse(readFileSync(join(out.location, 'manifest.json'), 'utf8'));
    expect(manifest.files.map((f: { name: string }) => f.name)).toEqual(
      expect.arrayContaining(saved.map((a) => a.path)),
    );
    expect(manifest.counts).toMatchObject({ attachments: 3, attachments_skipped: 4 });
  });

  it('leaves attachments alone when not asked for', async () => {
    const s = stub();
    const out = parseToolResult<Result>(await call('outlook_export_messages', s.client, {}));
    expect(out.counts).toEqual({ messages: 2, attachments: 0, attachments_skipped: 0 });
    expect(s.calls.some((c) => c.path.includes('/attachments'))).toBe(false);
  });

  it('saves event attachments from /me/events/{id}/attachments', async () => {
    const s = stub();
    const out = parseToolResult<Result>(
      await call('outlook_export_events', s.client, {
        since: '2026-03-01T00:00:00Z',
        until: '2026-04-01T00:00:00Z',
        attachments: true,
      }),
    );
    expect(out.counts).toEqual({ events: 1, attachments: 1, attachments_skipped: 0 });
    expect(s.calls.some((c) => c.path === '/me/events/ev-1/attachments')).toBe(true);
    const part = out.files.find((f) => f.path?.endsWith('part-0001.jsonl'))!;
    const [rec] = records(part.path!);
    const [a] = rec.attachments as Entry[];
    expect(readFileSync(join(out.location, a.path!), 'utf8')).toBe('agenda');
  });

  it('puts the attachments in the zip in blob mode', async () => {
    const fetchImpl = vi.fn(async () => new Response(null, { status: 200 })) as unknown as typeof fetch;
    const env = { MCP_BLOB_BASE_URL: 'https://host.test/b/reg123', MCP_BLOB_SIGNING_KEY: 'test-key' };
    const out = parseToolResult<Result>(
      await call('outlook_export_messages', stub().client, { attachments: true }, { env, fetchImpl }),
    );
    expect(out.mode).toBe('blob');
    expect(out.files.filter((f) => f.name?.startsWith('attachments/'))).toHaveLength(3);

    // Open the uploaded zip: the attachment bytes are there, already-compressed
    // media is STOREd and the JSONL part is DEFLATEd.
    const put = (fetchImpl as unknown as { mock: { calls: [string, RequestInit][] } }).mock.calls[0][1];
    const zip = unzip(Buffer.from(put.body as ArrayBuffer));
    const pdfs = [...zip.entries()].filter(([n]) => n.endsWith('.pdf'));
    expect(pdfs).toHaveLength(2);
    expect(pdfs.map(([, e]) => e.method)).toEqual([0, 0]);
    expect(pdfs.map(([, e]) => e.data.toString('utf8')).sort()).toEqual(['%PDF-fake', 'second']);
    expect(zip.get('part-0001.jsonl')?.method).toBe(8);
    expect(zip.get('part-0001.jsonl')?.data.toString('utf8')).toContain('"id":"id-1"');
  });
});

/** Read a zip's entries from its central directory (STORE and DEFLATE only). */
function unzip(buf: Buffer): Map<string, { method: number; data: Buffer }> {
  const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const count = buf.readUInt16LE(eocd + 10);
  let at = buf.readUInt32LE(eocd + 16);
  const out = new Map<string, { method: number; data: Buffer }>();
  for (let i = 0; i < count; i++) {
    const method = buf.readUInt16LE(at + 10);
    const size = buf.readUInt32LE(at + 20);
    const nameLen = buf.readUInt16LE(at + 28);
    const extraLen = buf.readUInt16LE(at + 30);
    const commentLen = buf.readUInt16LE(at + 32);
    const local = buf.readUInt32LE(at + 42);
    const name = buf.toString('utf8', at + 46, at + 46 + nameLen);
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const raw = buf.subarray(start, start + size);
    out.set(name, { method, data: method === 8 ? inflateRawSync(raw) : Buffer.from(raw) });
    at += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/**
 * A mailbox that honours the keyset `$filter` and `$top`, where every message
 * has one file attachment. `onDownload` runs before each attachment download.
 */
function keysetStub(n: number, onDownload: (id: string) => void = () => {}) {
  const mailbox = Array.from({ length: n }, (_, i) => ({
    Id: `m-${i + 1}`,
    Subject: `Subject ${i + 1}`,
    ReceivedDateTime: new Date(Date.UTC(2026, 0, 1, 0, i + 1)).toISOString(),
    HasAttachments: true,
    Body: { ContentType: 'Text', Content: 'body' },
  }));
  const att = (id: string) => ({
    '@odata.type': FILE,
    Id: `att-${id}`,
    Name: `${id}.txt`,
    ContentType: 'text/plain',
    Size: 8,
    ContentBytes: b64(`bytes-${id}`),
  });
  const calls: { path: string; opts?: Opts }[] = [];
  const get = vi.fn(async (path: string, opts?: Opts) => {
    calls.push({ path, opts });
    if (path === '/me') return { EmailAddress: 'me@example.test' };
    const one = /^\/me\/messages\/([^/]+)\/attachments\/([^/]+)$/.exec(path);
    if (one) {
      const id = decodeURIComponent(one[1]);
      onDownload(id);
      return att(id);
    }
    const list = /^\/me\/messages\/([^/]+)\/attachments$/.exec(path);
    if (list) {
      const { ContentBytes: _b, ...meta } = att(decodeURIComponent(list[1]));
      return { value: [meta] };
    }
    const ge = /ReceivedDateTime ge (\S+)/.exec(String(opts?.query?.$filter ?? ''))?.[1];
    const top = Number(opts?.query?.$top ?? 1000);
    const rows = mailbox.filter((m) => !ge || Date.parse(m.ReceivedDateTime) >= Date.parse(ge));
    return { value: rows.slice(0, top), ...(rows.length > top ? { '@odata.nextLink': 'https://next.test' } : {}) };
  });
  const downloads = () => calls.filter((c) => /\/attachments\/[^/]+$/.test(c.path)).map((c) => c.path);
  return { client: { get, getAbsolute: vi.fn(), write: vi.fn() } as unknown as OutlookClient, calls, downloads };
}

const partRecords = (out: Result, part: number) =>
  records(out.files.find((f) => f.path?.endsWith(`part-000${part}.jsonl`))!.path!);

describe('export attachments: stopping mid-page', () => {
  it('stops on a credential failure during an attachment download and loses nothing on resume', async () => {
    let refuse = true;
    const s = keysetStub(3, (id) => {
      if (id === 'm-2' && refuse) {
        throw new McpToolError('Outlook rejected the access token (401).', { status: 401, kind: 'credential_rejected' });
      }
    });
    const first = parseToolResult<Result>(await call('outlook_export_messages', s.client, { attachments: true }, { pageSize: 1 }));
    expect(first.has_more).toBe(true);
    expect(first.cursor).toBeTruthy();
    expect(first.errors?.join(' ')).toMatch(/credential/i);
    // m-1 is kept; m-2 (whose attachment hit the 401) is NOT written with a skipped attachment.
    expect(first.counts).toEqual({ messages: 1, attachments: 1, attachments_skipped: 0 });
    expect(partRecords(first, 1).map((r) => r.id)).toEqual(['m-1']);
    // The refused credential is not retried for every remaining attachment.
    expect(s.downloads().filter((p) => p.includes('m-2'))).toHaveLength(1);
    expect(s.downloads().some((p) => p.includes('m-3'))).toBe(false);

    refuse = false;
    const second = parseToolResult<Result>(
      await call('outlook_export_messages', s.client, { cursor: first.cursor }, { pageSize: 1 }),
    );
    expect(second.has_more).toBe(false);
    expect(second.counts).toEqual({ messages: 3, attachments: 3, attachments_skipped: 0 });
    const resumed = partRecords(second, 2);
    expect(resumed.map((r) => r.id)).toEqual(['m-2', 'm-3']);
    const [a] = resumed[0].attachments as Entry[];
    expect(a.skipped).toBeUndefined();
    expect(readFileSync(join(second.location, a.path!), 'utf8')).toBe('bytes-m-2');

    // On disk the manifest is cumulative: every part's files, not just the latest part's.
    const manifest = JSON.parse(readFileSync(join(second.location, 'manifest.json'), 'utf8'));
    const names = manifest.files.map((f: { name: string }) => f.name);
    expect(names).toEqual(expect.arrayContaining(['part-0001.jsonl', 'part-0002.jsonl']));
    expect(names.filter((n: string) => n.startsWith('attachments/'))).toHaveLength(3);
    expect(new Set(names).size).toBe(names.length);
  });

  it('stops at budget_seconds inside a page of attachments, at a record boundary, and resumes there', async () => {
    let now = 0;
    // Each attachment download takes 20 s of the (fake) clock.
    const s = keysetStub(10, () => {
      now += 20_000;
    });
    const first = parseToolResult<Result>(
      await call('outlook_export_messages', s.client, { attachments: true, budget_seconds: 30 }, { clock: () => now }),
    );
    expect(first.has_more).toBe(true);
    expect(first.counts.messages).toBeGreaterThan(0);
    expect(first.counts.messages).toBeLessThan(10);
    expect(first.counts.attachments_skipped).toBe(0);
    // Work stopped soon after the budget: not every attachment in the page was fetched.
    expect(s.downloads().length).toBe(first.counts.messages);
    expect(s.downloads().length).toBeLessThanOrEqual(4);
    // With attachments on, a page asks for few rows, so one page cannot run unbounded.
    const listing = s.calls.find((c) => c.opts?.query?.$filter !== undefined)!;
    expect(Number(listing.opts!.query!.$top)).toBeLessThanOrEqual(50);

    let out = first;
    const ids = partRecords(first, 1).map((r) => r.id);
    for (let part = 2; out.has_more && part < 9; part++) {
      out = parseToolResult<Result>(
        await call('outlook_export_messages', s.client, { cursor: out.cursor, budget_seconds: 30 }, { clock: () => now }),
      );
      ids.push(...partRecords(out, part).map((r) => r.id));
    }
    expect(out.has_more).toBe(false);
    expect(ids).toEqual(Array.from({ length: 10 }, (_, i) => `m-${i + 1}`));
    expect(out.counts).toEqual({ messages: 10, attachments: 10, attachments_skipped: 0 });
  });
});
