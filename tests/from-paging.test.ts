import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { buildQueryString } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import { decodeFrom, encodeFrom } from '../src/tools/_paging.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `from` / `next_from`: oldest-first keyset paging from a UTC time, shared by
 * outlook_get_unread and outlook_list_messages. Fixtures are synthetic.
 */

const wire = (path: string, opts?: { query?: Record<string, unknown> }) =>
  `${path}${buildQueryString((opts?.query ?? {}) as Record<string, string | number | undefined>)}`;

const msg = (id: string, at: string) => ({
  Id: id,
  Subject: `Subject ${id}`,
  ReceivedDateTime: at,
  IsRead: false,
});

/** A client whose folder listing returns `list` (oldest first, as the API would for `asc`). */
function stub(list: Record<string, unknown>[], nextLink?: string) {
  const paths: string[] = [];
  const get = vi.fn(async (p: string, o?: { query?: Record<string, unknown> }) => {
    paths.push(decodeURIComponent(wire(p, o)));
    if (p.includes('/mailfolders/')) return { value: list, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) };
    return { Id: 'x', Body: { ContentType: 'Text', Content: 'body' } };
  });
  const getAbsolute = vi.fn(async () => ({ value: [] }));
  const client = { get, getAbsolute, write: vi.fn() } as unknown as OutlookClient;
  return { client, paths, getAbsolute };
}

type Out = {
  count: number;
  items: Record<string, unknown>[];
  next_from?: string;
  has_more?: boolean;
  nextLink?: string;
};

async function call(client: OutlookClient, tool: string, args: Record<string, unknown>) {
  const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
  const res = await h.callTool(tool, args);
  await h.close();
  return res;
}

const ids = (out: Out) => out.items.map((i) => i.id ?? i.Id);
const token = (t: string, tied: string[] = []) =>
  Buffer.from(JSON.stringify({ t, ids: tied }), 'utf8').toString('base64url');

describe('next_from token', () => {
  it('round-trips a position', () => {
    const pos = { t: '2026-10-01T08:00:00Z', ids: ['a', 'b'] };
    expect(decodeFrom(encodeFrom(pos))).toEqual(pos);
  });

  it('is base64url JSON of { t, ids }', () => {
    const tok = encodeFrom({ t: '2026-10-01T08:00:00Z', ids: ['a'] });
    expect(tok).toMatch(/^[A-Za-z0-9_-]+$/);
    expect(JSON.parse(Buffer.from(tok, 'base64url').toString('utf8'))).toEqual({
      t: '2026-10-01T08:00:00Z',
      ids: ['a'],
    });
  });

  it('accepts an ISO 8601 UTC time as a position with no ids', () => {
    expect(decodeFrom('2026-10-01T00:00:00Z')).toEqual({ t: '2026-10-01T00:00:00Z', ids: [] });
  });

  it('normalises an offset time to UTC', () => {
    expect(decodeFrom('2026-10-01T02:00:00+02:00')).toEqual({ t: '2026-10-01T00:00:00Z', ids: [] });
  });

  it('reads a time with no zone (or a bare date) as UTC, never local time', () => {
    expect(decodeFrom('2026-10-01T08:30:00')).toEqual({ t: '2026-10-01T08:30:00Z', ids: [] });
    expect(decodeFrom('2026-10-01T08:30')).toEqual({ t: '2026-10-01T08:30:00Z', ids: [] });
    expect(decodeFrom('2026-10-01')).toEqual({ t: '2026-10-01T00:00:00Z', ids: [] });
  });

  it.each(['yesterday', 'not-a-token!', '2026-02-30', '2026-13-45T00:00:00Z', token('nope'), 'e30'])(
    'rejects %s with a hint',
    (bad) => {
      expect(() => decodeFrom(bad)).toThrow(/from/i);
      try {
        decodeFrom(bad);
      } catch (e) {
        expect((e as { hint?: string }).hint).toMatch(/next_from unchanged, or an ISO 8601 time/);
      }
    },
  );
});

describe.each([
  ['outlook_get_unread', { includeBody: false }],
  ['outlook_list_messages', {}],
])('%s with from', (tool, base) => {
  it('queries ReceivedDateTime ge <t> first, ascending, with $top = limit', async () => {
    const s = stub([]);
    await call(s.client, tool, { ...base, from: '2026-10-01T00:00:00Z', limit: 10 });
    const p = s.paths[0];
    expect(p).toMatch(/\$filter=ReceivedDateTime ge 2026-10-01T00:00:00Z( and IsRead eq false)?(&|$)/);
    expect(p).toContain('$orderby=ReceivedDateTime asc');
    expect(p).toContain('$top=10');
  });

  it('returns items oldest first with has_more false and a resumable next_from', async () => {
    const s = stub([msg('a', '2026-10-01T01:00:00Z'), msg('b', '2026-10-01T02:00:00Z')]);
    const out = parseToolResult<Out>(await call(s.client, tool, { ...base, from: '2026-10-01T00:00:00Z' }));
    expect(ids(out)).toEqual(['a', 'b']);
    expect(out.has_more).toBe(false);
    expect(out).not.toHaveProperty('nextLink');
    expect(decodeFrom(out.next_from!)).toEqual({ t: '2026-10-01T02:00:00Z', ids: ['b'] });
  });

  it('asks for the tied ids on top of limit, skips them, and keeps ties in the next token', async () => {
    // Page 1 ended at 02:00 with `b`; `c` shares that second and must not be lost.
    const s = stub([
      msg('b', '2026-10-01T02:00:00Z'),
      msg('c', '2026-10-01T02:00:00Z'),
      msg('d', '2026-10-01T03:00:00Z'),
    ]);
    const out = parseToolResult<Out>(
      await call(s.client, tool, { ...base, from: token('2026-10-01T02:00:00Z', ['b']), limit: 1 }),
    );
    expect(s.paths[0]).toContain('$filter=ReceivedDateTime ge 2026-10-01T02:00:00Z');
    expect(s.paths[0]).toContain('$top=2');
    expect(ids(out)).toEqual(['c']);
    expect(out.has_more).toBe(true);
    // `t` did not move, so the previous tie carries forward with the new one.
    expect(decodeFrom(out.next_from!)).toEqual({ t: '2026-10-01T02:00:00Z', ids: ['b', 'c'] });
  });

  it('collects every returned id sharing the newest timestamp', async () => {
    const s = stub([
      msg('a', '2026-10-01T01:00:00Z'),
      msg('b', '2026-10-01T02:00:00Z'),
      msg('c', '2026-10-01T02:00:00Z'),
    ]);
    const out = parseToolResult<Out>(await call(s.client, tool, { ...base, from: '2026-10-01T00:00:00Z' }));
    expect(decodeFrom(out.next_from!)).toEqual({ t: '2026-10-01T02:00:00Z', ids: ['b', 'c'] });
  });

  it('reports has_more when the API has another page', async () => {
    const s = stub([msg('a', '2026-10-01T01:00:00Z')], 'https://outlook.office.com/api/v2.0/next');
    const out = parseToolResult<Out>(await call(s.client, tool, { ...base, from: '2026-10-01T00:00:00Z', limit: 1 }));
    expect(out.has_more).toBe(true);
    expect(out).not.toHaveProperty('nextLink');
  });

  it('returns the input position as next_from on an empty page', async () => {
    const s = stub([]);
    const from = token('2026-10-01T02:00:00Z', ['b']);
    const out = parseToolResult<Out>(await call(s.client, tool, { ...base, from }));
    expect(out.count).toBe(0);
    expect(out.has_more).toBe(false);
    expect(decodeFrom(out.next_from!)).toEqual({ t: '2026-10-01T02:00:00Z', ids: ['b'] });
  });

  it('returns the input time as next_from on an empty page from an ISO time', async () => {
    const s = stub([]);
    const out = parseToolResult<Out>(await call(s.client, tool, { ...base, from: '2026-10-01T00:00:00Z' }));
    expect(decodeFrom(out.next_from!)).toEqual({ t: '2026-10-01T00:00:00Z', ids: [] });
  });

  it('rejects a bad from before any request', async () => {
    const s = stub([]);
    const res = await call(s.client, tool, { ...base, from: 'last tuesday' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/next_from unchanged, or an ISO 8601 time/);
    expect(s.paths).toHaveLength(0);
  });

  it('rejects from with nextLink', async () => {
    const s = stub([]);
    const res = await call(s.client, tool, {
      ...base,
      from: '2026-10-01T00:00:00Z',
      nextLink: 'https://outlook.office.com/api/v2.0/next',
    });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/next_from/);
    expect(s.paths).toHaveLength(0);
    expect(s.getAbsolute).not.toHaveBeenCalled();
  });
});

describe('outlook_get_unread with from', () => {
  it('keeps the unread filter after the date bound', async () => {
    const s = stub([]);
    await call(s.client, 'outlook_get_unread', { from: '2026-10-01T00:00:00Z', includeBody: false });
    expect(s.paths[0]).toContain('$filter=ReceivedDateTime ge 2026-10-01T00:00:00Z and IsRead eq false');
  });

  it('rejects from with sinceHours', async () => {
    const s = stub([]);
    const res = await call(s.client, 'outlook_get_unread', { from: '2026-10-01T00:00:00Z', sinceHours: 24 });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/sinceHours/);
    expect(s.paths).toHaveLength(0);
  });

  it('fetches bodies only for the kept rows', async () => {
    const s = stub([msg('b', '2026-10-01T02:00:00Z'), msg('c', '2026-10-01T03:00:00Z')]);
    const out = parseToolResult<Out>(
      await call(s.client, 'outlook_get_unread', { from: token('2026-10-01T02:00:00Z', ['b']) }),
    );
    expect(ids(out)).toEqual(['c']);
    expect(out.items[0].body).toBe('body');
    expect(s.paths.filter((p) => p.startsWith('/me/messages/'))).toHaveLength(1);
  });
});

describe('outlook_list_messages with from', () => {
  it('adds the unread filter after the date bound when unreadOnly', async () => {
    const s = stub([]);
    await call(s.client, 'outlook_list_messages', { from: '2026-10-01T00:00:00Z', unreadOnly: true, folder: 'archive' });
    expect(s.paths[0]).toContain('/me/mailfolders/archive/messages');
    expect(s.paths[0]).toContain('$filter=ReceivedDateTime ge 2026-10-01T00:00:00Z and IsRead eq false');
  });

  it('has no unread filter without unreadOnly', async () => {
    const s = stub([]);
    await call(s.client, 'outlook_list_messages', { from: '2026-10-01T00:00:00Z' });
    expect(s.paths[0]).not.toContain('IsRead eq false');
  });

  it.each([
    ['search', { search: 'report' }],
    ['skip', { skip: 25 }],
  ])('rejects from with %s', async (name, extra) => {
    const s = stub([]);
    const res = await call(s.client, 'outlook_list_messages', { from: '2026-10-01T00:00:00Z', ...extra });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toContain(name);
    expect(s.paths).toHaveLength(0);
  });

  it('carries next_from and has_more in the raw view too, without the API link', async () => {
    const s = stub([msg('a', '2026-10-01T01:00:00Z')], 'https://outlook.office.com/api/v2.0/next');
    const out = parseToolResult<Record<string, unknown>>(
      await call(s.client, 'outlook_list_messages', { from: '2026-10-01T00:00:00Z', view: 'raw' }),
    );
    expect(out).not.toHaveProperty('@odata.nextLink');
    expect(out.has_more).toBe(true);
    expect(decodeFrom(out.next_from as string)).toEqual({ t: '2026-10-01T01:00:00Z', ids: ['a'] });
  });

  it('without from stays newest first with nextLink and no next_from', async () => {
    const s = stub([msg('a', '2026-10-01T01:00:00Z')], 'https://outlook.office.com/api/v2.0/next');
    const out = parseToolResult<Out>(await call(s.client, 'outlook_list_messages', {}));
    expect(s.paths[0]).toContain('$orderby=ReceivedDateTime desc');
    expect(out.nextLink).toBe('https://outlook.office.com/api/v2.0/next');
    expect(out).not.toHaveProperty('next_from');
    expect(out).not.toHaveProperty('has_more');
  });
});

describe('outlook_get_unread without from', () => {
  it('stays newest first with nextLink and no next_from', async () => {
    const s = stub([msg('a', '2026-10-01T01:00:00Z')], 'https://outlook.office.com/api/v2.0/next');
    const out = parseToolResult<Out>(await call(s.client, 'outlook_get_unread', { includeBody: false }));
    expect(s.paths[0]).toContain('$orderby=ReceivedDateTime desc');
    expect(out.nextLink).toBe('https://outlook.office.com/api/v2.0/next');
    expect(out).not.toHaveProperty('next_from');
    expect(out).not.toHaveProperty('has_more');
  });
});
