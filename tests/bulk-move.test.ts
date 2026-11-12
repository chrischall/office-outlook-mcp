import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { buildMoveFilter, registerBulkTools } from '../src/tools/bulk.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_move_messages` — move every message in a folder that matches a
 * filter, under one confirmation. The client is an in-memory folder that pages
 * its `$filter`ed listing; fixtures are synthetic (example.com senders).
 */

type Row = { Id: string; Subject?: string; From?: { EmailAddress?: { Address?: string; Name?: string } }; ReceivedDateTime?: string; IsRead?: boolean };
type Call = { method: string; path: string; body?: unknown; query?: Record<string, unknown> };

const BASE = 'https://outlook.office.com/api/v2.0';

function msg(i: number, over: Partial<Row> = {}): Row {
  return {
    Id: `m-${i}`,
    Subject: `Newsletter ${i}`,
    From: { EmailAddress: { Address: 'news@example.com', Name: 'News' } },
    ReceivedDateTime: `2026-06-${String((i % 28) + 1).padStart(2, '0')}T10:00:00Z`,
    IsRead: false,
    ...over,
  };
}

/**
 * A folder of `rows`, served in pages of `pageSize` (default 100) with an
 * absolute `@odata.nextLink` — the client does not evaluate `$filter`, the
 * test asserts on the one it was sent. A move POST returns a new id.
 */
function stub(
  rows: Row[],
  opts: {
    pageSize?: number;
    moveFail?: Record<string, Error>;
    patchFail?: string[];
    moveDelayMs?: number;
    onMove?: (n: number) => void;
    moveNoId?: string[];
    /** Folder id → DisplayName served by GET /me/mailfolders/{id}; others 404. */
    folderNames?: Record<string, string>;
  } = {},
) {
  let folder = [...rows];
  const calls: Call[] = [];
  const pageSize = opts.pageSize ?? 100;
  let moves = 0;
  const page = (skip: number) => {
    const value = folder.slice(skip, skip + pageSize);
    const more = skip + pageSize < folder.length;
    return { value, ...(more ? { '@odata.nextLink': `${BASE}/me/mailfolders/inbox/messages?$skip=${skip + pageSize}` } : {}) };
  };
  const get = vi.fn(async (path: string, o: { query?: Record<string, unknown> } = {}) => {
    calls.push({ method: 'GET', path, query: o.query });
    if (/^\/me\/mailfolders\/[^/]+\/messages$/.test(path)) return page(0);
    const folderMatch = /^\/me\/mailfolders\/([^/]+)$/.exec(path);
    if (folderMatch) {
      const name = opts.folderNames?.[decodeURIComponent(folderMatch[1]!)];
      if (name === undefined) throw new ApiError(404, 'Folder not found');
      return { Id: `id-${folderMatch[1]}`, DisplayName: name };
    }
    throw new Error(`unexpected GET ${path}`);
  });
  const getAbsolute = vi.fn(async (url: string) => {
    calls.push({ method: 'GETABS', path: url });
    const skip = Number(new URL(url).searchParams.get('$skip'));
    return page(skip);
  });
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    const id = decodeURIComponent(path.split('/')[3] ?? '');
    if (method === 'POST' && path.endsWith('/move')) {
      if (opts.moveDelayMs) await new Promise((r) => setTimeout(r, opts.moveDelayMs));
      const err = opts.moveFail?.[id];
      if (err) throw err;
      moves++;
      opts.onMove?.(moves);
      folder = folder.filter((r) => r.Id !== id);
      if (opts.moveNoId?.includes(id)) return {};
      return { Id: `new-${id}`, ParentFolderId: (body as { DestinationId?: string }).DestinationId };
    }
    if (method === 'PATCH') {
      if (opts.patchFail?.includes(id)) throw new ApiError(500, 'Outlook 500');
      return { Id: id, IsRead: true };
    }
    throw new Error(`unexpected ${method} ${path}`);
  });
  const client = { get, getAbsolute, write } as unknown as OutlookClient;
  return {
    client,
    calls,
    moves: () => calls.filter((c) => c.method === 'POST'),
    patches: () => calls.filter((c) => c.method === 'PATCH'),
    firstQuery: () => calls.find((c) => c.method === 'GET')?.query ?? {},
    /** Replace the folder's contents (mail arriving between preview and confirm). */
    setFolder: (next: Row[]) => {
      folder = [...next];
    },
  };
}

const harness = (client: OutlookClient) => createTestHarness((s: McpServer) => registerBulkTools(s, client));

type Preview = {
  status?: string;
  error?: string;
  confirmToken?: string;
  preview?: {
    action?: string;
    matched?: number;
    willMove?: number;
    beyondMax?: number;
    destination?: string;
    destinationId?: string;
    source?: string;
    sourceId?: string;
    markRead?: boolean;
    filter?: Record<string, unknown>;
    sample?: { from?: string; subject?: string; received?: string }[];
  };
};
type Moved = {
  moved: number;
  failed: { id: string; subject?: string; error: string }[];
  failedCount: number;
  markedRead?: number;
  remaining: number;
  destination: string;
  destinationId?: string;
  unknown?: { count: number; ids: string[]; warning: string };
  matched?: number;
};

async function confirmed(h: TestHarness, args: Record<string, unknown>) {
  const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
  expect(first.status).toBe('confirmation-required');
  return h.callTool('outlook_move_messages', { ...args, confirmToken: first.confirmToken });
}

const ENV_KEYS = ['MCP_CONFIRM_MODE', 'MCP_CONFIRM_TTL_SECONDS', 'MCP_CONFIRM_SECRET'] as const;
let savedEnv: Record<string, string | undefined>;
beforeEach(() => {
  savedEnv = Object.fromEntries(ENV_KEYS.map((k) => [k, process.env[k]]));
  for (const k of ENV_KEYS) delete process.env[k];
});
afterEach(() => {
  for (const k of ENV_KEYS) {
    if (savedEnv[k] === undefined) delete process.env[k];
    else process.env[k] = savedEnv[k];
  }
});

describe('buildMoveFilter', () => {
  it('OR-joins senders and ANDs the date and read conditions', () => {
    expect(
      buildMoveFilter({
        from: ['a@example.com', 'b@example.com'],
        receivedBefore: '2026-08-01T00:00:00Z',
        receivedAfter: '2026-01-01',
        isRead: true,
      }),
    ).toBe(
      "(From/EmailAddress/Address eq 'a@example.com' or From/EmailAddress/Address eq 'b@example.com')" +
        ' and ReceivedDateTime lt 2026-08-01T00:00:00Z and ReceivedDateTime ge 2026-01-01T00:00:00Z and IsRead eq true',
    );
  });

  it("escapes a quote in an address by doubling it", () => {
    expect(buildMoveFilter({ from: ["o'brien@example.com"] })).toBe(
      "(From/EmailAddress/Address eq 'o''brien@example.com')",
    );
  });

  it('reads a time with no zone as UTC and keeps a real offset', () => {
    expect(buildMoveFilter({ receivedBefore: '2026-08-01T09:30:00' })).toBe('ReceivedDateTime lt 2026-08-01T09:30:00Z');
    expect(buildMoveFilter({ receivedBefore: '2026-08-01T09:30:00-04:00' })).toBe(
      'ReceivedDateTime lt 2026-08-01T13:30:00Z',
    );
  });

  it('refuses a date that is not ISO 8601', () => {
    expect(() => buildMoveFilter({ receivedBefore: 'last tuesday' })).toThrow(/receivedBefore/);
  });

  it('leaves the subject conditions out of $filter (they are applied client-side)', () => {
    expect(buildMoveFilter({ subjectContains: ['invoice'] })).toBeUndefined();
  });
});

describe('outlook_move_messages', () => {
  it('refuses an unfiltered "move everything" and reads nothing', async () => {
    const s = stub([msg(1)]);
    const h = await harness(s.client);
    const res = await h.callTool('outlook_move_messages', { destination: 'archive' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/at least one/i);
    // exceptSubjectContains / markRead alone narrow nothing.
    const res2 = await h.callTool('outlook_move_messages', {
      destination: 'archive',
      exceptSubjectContains: ['keep'],
      markRead: true,
    });
    expect(res2.isError).toBe(true);
    expect(s.calls).toHaveLength(0);
    await h.close();
  });

  it('previews once with the count, the filter and a 10-message sample, writing nothing', async () => {
    const s = stub(Array.from({ length: 15 }, (_, i) => msg(i)));
    const h = await harness(s.client);
    const first = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { from: ['News@Example.com'], destination: 'archive' }),
    );
    expect(first.status).toBe('confirmation-required');
    expect(first.preview?.matched).toBe(15);
    expect(first.preview?.willMove).toBe(15);
    expect(first.preview?.destination).toBe('archive');
    expect(first.preview?.filter).toMatchObject({ folder: 'inbox', from: ['News@Example.com'] });
    expect(first.preview?.sample).toHaveLength(10);
    expect(first.preview?.sample?.[0]).toEqual({
      from: 'news@example.com',
      subject: 'Newsletter 0',
      received: '2026-06-01T10:00:00Z',
    });
    expect(s.moves()).toHaveLength(0);
    // The listing asked Outlook to do the sender filtering, with a narrow $select.
    const q = s.firstQuery();
    expect(q.$filter).toBe("(From/EmailAddress/Address eq 'News@Example.com')");
    expect(q.$select).toBe('Id,Subject,From,ReceivedDateTime,IsRead');
    expect(q.$top).toBe(100);
    expect(s.calls[0]?.path).toBe('/me/mailfolders/inbox/messages');
    await h.close();
  });

  it('pages across nextLink and moves every match after confirmation', async () => {
    const s = stub(Array.from({ length: 250 }, (_, i) => msg(i)), { pageSize: 100 });
    const h = await harness(s.client);
    const res = parseToolResult<Moved>(await confirmed(h, { from: ['news@example.com'], destination: 'archive' }));
    expect(res.moved).toBe(250);
    expect(res.failedCount).toBe(0);
    expect(res.remaining).toBe(0);
    expect(res.destination).toBe('archive');
    expect(s.moves()).toHaveLength(250);
    expect(s.moves()[0]).toMatchObject({ path: '/me/messages/m-0/move', body: { DestinationId: 'archive' } });
    await h.close();
  });

  it('applies subjectContains (any of) and exceptSubjectContains (none of) case-insensitively', async () => {
    const rows = [
      msg(1, { Subject: 'Your INVOICE for June' }),
      msg(2, { Subject: 'Receipt #42' }),
      msg(3, { Subject: 'Invoice — DO NOT ARCHIVE' }),
      msg(4, { Subject: 'Weekly digest' }),
      msg(5, { Subject: undefined }),
    ];
    const s = stub(rows);
    const h = await harness(s.client);
    const args = {
      subjectContains: ['invoice', 'receipt'],
      exceptSubjectContains: ['do not archive'],
      destination: 'archive',
    };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    expect(first.preview?.matched).toBe(2);
    expect(first.preview?.sample?.map((x) => x.subject)).toEqual(['Your INVOICE for June', 'Receipt #42']);
    // No server-side condition: the folder is listed with no $filter at all.
    expect(s.firstQuery().$filter).toBeUndefined();
    const res = parseToolResult<Moved>(
      await h.callTool('outlook_move_messages', { ...args, confirmToken: first.confirmToken }),
    );
    expect(res.moved).toBe(2);
    expect(s.moves().map((c) => c.path)).toEqual(['/me/messages/m-1/move', '/me/messages/m-2/move']);
    await h.close();
  });

  it('double-checks the sender client-side, ignoring case', async () => {
    const rows = [msg(1), msg(2, { From: { EmailAddress: { Address: 'someone@example.com' } } })];
    const s = stub(rows);
    const h = await harness(s.client);
    const first = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { from: ['NEWS@example.com'], destination: 'archive' }),
    );
    expect(first.preview?.matched).toBe(1);
    await h.close();
  });

  it('caps at max: says "more than max" in the preview, moves max, and reports the rest as remaining', async () => {
    const s = stub(Array.from({ length: 12 }, (_, i) => msg(i)));
    const h = await harness(s.client);
    const args = { from: ['news@example.com'], destination: 'archive', max: 5 };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    expect(first.preview?.matched).toBe(12);
    expect(first.preview?.willMove).toBe(5);
    expect(first.preview?.beyondMax).toBe(7);
    expect(JSON.stringify(first.preview)).toMatch(/more than 5/);
    const res = parseToolResult<Moved>(
      await h.callTool('outlook_move_messages', { ...args, confirmToken: first.confirmToken }),
    );
    expect(res.moved).toBe(5);
    expect(res.remaining).toBe(7);
    expect(s.moves()).toHaveLength(5);
    await h.close();
  });

  it('rejects a max above the hard limit of 5000', async () => {
    const s = stub([msg(1)]);
    const h = await harness(s.client);
    const res = await h.callTool('outlook_move_messages', { from: ['news@example.com'], destination: 'archive', max: 5001 });
    expect(res.isError).toBe(true);
    await h.close();
  });

  it('re-previews instead of moving when the matched set changed between preview and confirm', async () => {
    const s = stub(Array.from({ length: 3 }, (_, i) => msg(i)));
    const h = await harness(s.client);
    const args = { from: ['news@example.com'], destination: 'archive' };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    expect(first.preview?.matched).toBe(3);
    // A fourth message from the sender arrives before the user confirms.
    s.setFolder(Array.from({ length: 4 }, (_, i) => msg(i)));
    const second = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { ...args, confirmToken: first.confirmToken }),
    );
    expect(second.error).toBe('DRAFT_CHANGED');
    expect(second.preview?.matched).toBe(4);
    expect(second.confirmToken).toEqual(expect.any(String));
    expect(s.moves()).toHaveLength(0);
    // The fresh token moves the set it was shown.
    const third = parseToolResult<Moved>(
      await h.callTool('outlook_move_messages', { ...args, confirmToken: second.confirmToken }),
    );
    expect(third.moved).toBe(4);
    await h.close();
  });

  it('binds markRead and the destination into the token', async () => {
    const s = stub([msg(1)]);
    const h = await harness(s.client);
    const args = { from: ['news@example.com'], destination: 'archive' };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    const swapped = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { ...args, markRead: true, confirmToken: first.confirmToken }),
    );
    expect(swapped.error).toBe('DRAFT_CHANGED');
    const elsewhere = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { ...args, destination: 'deleteditems', confirmToken: first.confirmToken }),
    );
    expect(elsewhere.error).toBe('DRAFT_CHANGED');
    expect(s.moves()).toHaveLength(0);
    await h.close();
  });

  it('marks each moved message read under its NEW id, skipping ones already read', async () => {
    const s = stub([msg(1), msg(2, { IsRead: true }), msg(3)]);
    const h = await harness(s.client);
    const res = parseToolResult<Moved>(
      await confirmed(h, { from: ['news@example.com'], destination: 'archive', markRead: true }),
    );
    expect(res.moved).toBe(3);
    expect(res.markedRead).toBe(2);
    expect(s.patches().map((c) => [c.path, c.body])).toEqual([
      ['/me/messages/new-m-1', { IsRead: true }],
      ['/me/messages/new-m-3', { IsRead: true }],
    ]);
    await h.close();
  });

  it('keeps going past a failed move and reports it with its subject', async () => {
    const s = stub([msg(1), msg(2), msg(3)], { moveFail: { 'm-2': new ApiError(404, 'Outlook 404: ErrorItemNotFound') } });
    const h = await harness(s.client);
    const raw = await confirmed(h, { from: ['news@example.com'], destination: 'archive' });
    expect(raw.isError).toBeFalsy();
    const res = parseToolResult<Moved>(raw);
    expect(res.moved).toBe(2);
    expect(res.failedCount).toBe(1);
    expect(res.failed[0]).toMatchObject({ id: 'm-2', subject: 'Newsletter 2' });
    expect(res.failed[0]?.error).toMatch(/404/);
    await h.close();
  });

  it('lists at most 20 failures but counts them all', async () => {
    const rows = Array.from({ length: 25 }, (_, i) => msg(i));
    const moveFail = Object.fromEntries(rows.map((r) => [r.Id, new ApiError(400, 'Outlook 400')]));
    const s = stub(rows, { moveFail });
    const h = await harness(s.client);
    const raw = await confirmed(h, { from: ['news@example.com'], destination: 'archive' });
    // Nothing moved and something failed: the one case that is an error.
    expect(raw.isError).toBe(true);
    const res = parseToolResult<Moved>(raw);
    expect(res.moved).toBe(0);
    expect(res.failedCount).toBe(25);
    expect(res.failed).toHaveLength(20);
    await h.close();
  });

  it('counts a timed-out move as unknown, never retries it, and carries on', async () => {
    const s = stub([msg(1), msg(2), msg(3)], {
      moveFail: { 'm-2': new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) },
    });
    const h = await harness(s.client);
    const raw = await confirmed(h, { from: ['news@example.com'], destination: 'archive', markRead: true });
    expect(raw.isError).toBeFalsy();
    const res = parseToolResult<Moved>(raw);
    expect(res.moved).toBe(2);
    expect(res.failedCount).toBe(0);
    expect(res.unknown?.count).toBe(1);
    expect(res.unknown?.ids).toEqual(['m-2']);
    expect(res.unknown?.warning).toMatch(/may already have happened/);
    expect(res.unknown?.warning).toMatch(/do not resend blindly/);
    // Sent once, not retried; and no mark-read for a message whose new id is unknown.
    expect(s.moves().filter((c) => c.path === '/me/messages/m-2/move')).toHaveLength(1);
    expect(s.patches()).toHaveLength(2);
    await h.close();
  });

  it('a failed mark-read does not undo or fail the move', async () => {
    const s = stub([msg(1), msg(2)], { patchFail: ['new-m-2'] });
    const h = await harness(s.client);
    const res = parseToolResult<Moved & { markReadFailed?: number }>(
      await confirmed(h, { from: ['news@example.com'], destination: 'archive', markRead: true }),
    );
    expect(res.moved).toBe(2);
    expect(res.markedRead).toBe(1);
    expect(res.markReadFailed).toBe(1);
    await h.close();
  });

  it('counts a move that returned no new id as a mark-read failure, not as marked', async () => {
    const s = stub([msg(1), msg(2)], { moveNoId: ['m-2'] });
    const h = await harness(s.client);
    const res = parseToolResult<Moved & { markReadFailed?: number }>(
      await confirmed(h, { from: ['news@example.com'], destination: 'archive', markRead: true }),
    );
    expect(res.moved).toBe(2);
    expect(res.markedRead).toBe(1);
    expect(res.markReadFailed).toBe(1);
    expect(s.patches()).toHaveLength(1);
    await h.close();
  });

  it('returns a plain "nothing matches" without asking to confirm', async () => {
    const s = stub([msg(1)]);
    const h = await harness(s.client);
    const raw = await h.callTool('outlook_move_messages', { from: ['nobody@example.com'], destination: 'archive' });
    expect(raw.isError).toBeFalsy();
    const res = parseToolResult<Moved & { status?: string }>(raw);
    expect(res.status).toBeUndefined();
    expect(res.moved).toBe(0);
    expect(res.matched).toBe(0);
    await h.close();
  });

  it('stops starting new moves once the caller cancels', async () => {
    const rows = Array.from({ length: 40 }, (_, i) => msg(i));
    const controller = new AbortController();
    const s = stub(rows, {
      moveDelayMs: 5,
      onMove: (n) => {
        if (n === 4) controller.abort();
      },
    });
    const h = await harness(s.client);
    const args = { from: ['news@example.com'], destination: 'archive' };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    await expect(
      h.client.callTool(
        { name: 'outlook_move_messages', arguments: { ...args, confirmToken: first.confirmToken } },
        { signal: controller.signal },
      ),
    ).rejects.toThrow();
    // Let anything still in flight settle.
    await new Promise((r) => setTimeout(r, 200));
    // At most the moves already in flight (concurrency 4) finish after the abort.
    expect(s.moves().length).toBeLessThanOrEqual(8);
    expect(s.moves().length).toBeGreaterThanOrEqual(4);
    await h.close();
  });

  it('defaults the source folder to inbox and takes another by id or name', async () => {
    const s = stub([msg(1)]);
    const h = await harness(s.client);
    await h.callTool('outlook_move_messages', { folder: 'Folder/1=', isRead: true, destination: 'archive' });
    expect(s.calls[0]?.path).toBe('/me/mailfolders/Folder%2F1%3D/messages');
    expect(s.firstQuery().$filter).toBe('IsRead eq true');
    await h.close();
  });
  it('names the source and destination folders in the preview and the result', async () => {
    const s = stub([msg(1), msg(2)], { folderNames: { inbox: 'Inbox', 'AAMk/alerts=': 'Alerts' } });
    const h = await harness(s.client);
    const args = { from: ['news@example.com'], destination: 'AAMk/alerts=' };
    const first = parseToolResult<Preview>(await h.callTool('outlook_move_messages', args));
    expect(first.preview?.action).toBe('Move 2 messages from Inbox to Alerts');
    expect(first.preview?.destination).toBe('Alerts');
    expect(first.preview?.destinationId).toBe('AAMk/alerts=');
    expect(first.preview?.source).toBe('Inbox');
    expect(first.preview?.sourceId).toBe('inbox');
    // The lookup asks for the name only, by the encoded id.
    const lookup = s.calls.find((c) => c.path === '/me/mailfolders/AAMk%2Falerts%3D');
    expect(lookup?.query).toEqual({ $select: 'DisplayName' });
    const res = parseToolResult<Moved>(
      await h.callTool('outlook_move_messages', { ...args, confirmToken: first.confirmToken }),
    );
    expect(res.moved).toBe(2);
    expect(res.destination).toBe('Alerts');
    expect(res.destinationId).toBe('AAMk/alerts=');
    // The move itself still targets the id.
    expect(s.moves()[0]).toMatchObject({ body: { DestinationId: 'AAMk/alerts=' } });
    await h.close();
  });

  it('resolves a well-known destination name the same way', async () => {
    const s = stub([msg(1)], { folderNames: { inbox: 'Inbox', archive: 'Archive' } });
    const h = await harness(s.client);
    const first = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { isRead: false, destination: 'archive' }),
    );
    expect(first.preview?.action).toBe('Move 1 message from Inbox to Archive');
    expect(first.preview?.destination).toBe('Archive');
    expect(first.preview?.destinationId).toBe('archive');
    await h.close();
  });

  it('falls back to the folder id when the name lookup fails', async () => {
    const s = stub([msg(1)], { folderNames: { inbox: 'Inbox' } });
    const h = await harness(s.client);
    const first = parseToolResult<Preview>(
      await h.callTool('outlook_move_messages', { isRead: false, destination: 'folder-xyz' }),
    );
    expect(first.status).toBe('confirmation-required');
    expect(first.preview?.action).toBe('Move 1 message from Inbox to folder-xyz');
    expect(first.preview?.destination).toBe('folder-xyz');
    expect(first.preview?.destinationId).toBe('folder-xyz');
    await h.close();
  });
});
