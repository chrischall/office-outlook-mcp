import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
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
  });
});
