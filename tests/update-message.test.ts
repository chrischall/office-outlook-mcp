import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerDirectoryTools } from '../src/tools/directory.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_update_message` — flag, categorise and mark read a batch of
 * messages — and `outlook_list_categories`. Fixtures are synthetic
 * (placeholder ids and category names).
 */

type Rec = Record<string, unknown>;

interface Stored {
  IsRead?: boolean;
  Flag?: { FlagStatus?: string };
  Categories?: string[];
}

/**
 * A mailbox of messages held in memory: GET returns the stored state, PATCH
 * merges into it and echoes the message back, as Outlook does. `ignore` makes
 * a PATCH succeed without changing anything; `fail` makes it throw.
 */
function stub(opts: { messages?: Record<string, Stored>; ignore?: string[]; fail?: string[]; echo?: boolean } = {}) {
  const store: Record<string, Stored> = structuredClone(
    opts.messages ?? { 'm-1': { IsRead: false, Categories: [] }, 'm-2': { IsRead: false, Categories: [] } },
  );
  const calls: { method: string; path: string; body?: unknown }[] = [];
  const idOf = (path: string) => decodeURIComponent(path.split('?')[0].split('/').pop() ?? '');
  const get = vi.fn(async (path: string) => {
    calls.push({ method: 'GET', path });
    const msg = store[idOf(path)];
    if (!msg) throw new Error('404 not found');
    return { Id: idOf(path), ...msg };
  });
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    const id = idOf(path);
    if (opts.fail?.includes(id)) throw new Error('Outlook said no');
    if (!opts.ignore?.includes(id)) Object.assign(store[id], body as Stored);
    return opts.echo === false ? undefined : { Id: id, ...store[id] };
  });
  const client = { get, write } as unknown as OutlookClient;
  const patches = () => calls.filter((c) => c.method === 'PATCH');
  return { client, calls, patches, store };
}

async function harness(client: OutlookClient) {
  return createTestHarness((s: McpServer) => registerWriteTools(s, client));
}

async function confirmed(h: TestHarness, args: Rec) {
  const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_update_message', args));
  return h.callTool('outlook_update_message', { ...args, confirmToken: first.confirmToken });
}

type Result = {
  updated?: number;
  failed?: number;
  results?: { id: string; ok: boolean; error?: string }[];
};

describe('outlook_update_message', () => {
  it('previews the batch and writes nothing without a token', async () => {
    const { client, patches } = stub();
    const h = await harness(client);
    const res = parseToolResult<{ status?: string }>(
      await h.callTool('outlook_update_message', { messageIds: ['m-1', 'm-2'], flag: 'flagged', isRead: true }),
    );
    expect(res.status).toBe('confirmation-required');
    const text = JSON.stringify(res);
    expect(text).toMatch(/2 messages/);
    expect(text).toMatch(/flagged/);
    expect(patches()).toHaveLength(0);
    await h.close();
  });

  it.each([
    ['flagged', 'Flagged'],
    ['complete', 'Complete'],
    ['none', 'NotFlagged'],
  ])('flag %s PATCHes FlagStatus %s on every id', async (flag, status) => {
    const { client, patches } = stub();
    const h = await harness(client);
    const res = parseToolResult<Result>(await confirmed(h, { messageIds: ['m-1', 'm-2'], flag }));
    expect(patches().map((p) => p.path).sort()).toEqual(['/me/messages/m-1', '/me/messages/m-2']);
    for (const p of patches()) expect(p.body).toEqual({ Flag: { FlagStatus: status } });
    expect(res.updated).toBe(2);
    expect(res.failed).toBe(0);
    expect(res.results).toEqual([
      { id: 'm-1', ok: true },
      { id: 'm-2', ok: true },
    ]);
    await h.close();
  });

  it('sends IsRead and replacing Categories in the same PATCH', async () => {
    const { client, patches, calls } = stub();
    const h = await harness(client);
    await confirmed(h, { messageIds: ['m-1'], isRead: true, categories: ['Blue category'] });
    expect(patches()).toEqual([
      { method: 'PATCH', path: '/me/messages/m-1', body: { IsRead: true, Categories: ['Blue category'] } },
    ]);
    // Replacing needs no read of the current list.
    expect(calls.filter((c) => c.method === 'GET')).toHaveLength(0);
    await h.close();
  });

  it('adds and removes categories against each message\'s current list', async () => {
    const { client, patches } = stub({
      messages: {
        'm-1': { Categories: ['Keep', 'Old'] },
        'm-2': { Categories: ['new'] },
      },
    });
    const h = await harness(client);
    const res = parseToolResult<Result>(
      await confirmed(h, { messageIds: ['m-1', 'm-2'], addCategories: ['New'], removeCategories: ['old'] }),
    );
    const byId = Object.fromEntries(patches().map((p) => [p.path, p.body]));
    // Removal matches case-insensitively; an add already present is not doubled.
    expect(byId['/me/messages/m-1']).toEqual({ Categories: ['Keep', 'New'] });
    expect(byId['/me/messages/m-2']).toEqual({ Categories: ['new'] });
    expect(res.updated).toBe(2);
    await h.close();
  });

  it('reports a failing id without abandoning the rest', async () => {
    const { client, patches } = stub({
      messages: { 'm-1': {}, 'm-2': {}, 'm-3': {} },
      fail: ['m-2'],
    });
    const h = await harness(client);
    const res = await confirmed(h, { messageIds: ['m-1', 'm-2', 'm-3'], isRead: true });
    expect(res.isError).toBeFalsy();
    const out = parseToolResult<Result>(res);
    expect(patches()).toHaveLength(3);
    expect(out.updated).toBe(2);
    expect(out.failed).toBe(1);
    expect(out.results?.[1]).toMatchObject({ id: 'm-2', ok: false, error: expect.stringMatching(/Outlook said no/) });
    expect(out.results?.[0]).toEqual({ id: 'm-1', ok: true });
    await h.close();
  });

  it('marks an id not ok when Outlook accepts the PATCH but the value did not change', async () => {
    const { client } = stub({ ignore: ['m-1'] });
    const h = await harness(client);
    const out = parseToolResult<Result>(await confirmed(h, { messageIds: ['m-1'], flag: 'complete' }));
    expect(out.updated).toBe(0);
    expect(out.results?.[0]).toMatchObject({ id: 'm-1', ok: false, error: expect.stringMatching(/Flag/) });
    await h.close();
  });

  it('re-reads to verify when the PATCH returns no body', async () => {
    const { client, calls } = stub({ echo: false });
    const h = await harness(client);
    const out = parseToolResult<Result>(await confirmed(h, { messageIds: ['m-1'], isRead: true }));
    expect(out.results?.[0]).toEqual({ id: 'm-1', ok: true });
    expect(calls.some((c) => c.method === 'GET' && c.path.includes('/me/messages/m-1'))).toBe(true);
    await h.close();
  });

  it('refuses a call with nothing to change', async () => {
    const { client, patches } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_message', { messageIds: ['m-1'] });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/Nothing to update/);
    expect(patches()).toHaveLength(0);
    await h.close();
  });

  it('refuses categories together with addCategories or removeCategories', async () => {
    const { client, patches } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_message', {
      messageIds: ['m-1'],
      categories: ['A'],
      addCategories: ['B'],
    });
    expect(res.isError).toBe(true);
    expect(patches()).toHaveLength(0);
    await h.close();
  });

  it('rejects an empty batch and one over 50 ids', async () => {
    const { client } = stub();
    const h = await harness(client);
    expect((await h.callTool('outlook_update_message', { messageIds: [], isRead: true })).isError).toBe(true);
    const many = Array.from({ length: 51 }, (_, i) => `m-${i}`);
    expect((await h.callTool('outlook_update_message', { messageIds: many, isRead: true })).isError).toBe(true);
    await h.close();
  });
});

describe('outlook_list_categories', () => {
  it('lists the master categories with their colours', async () => {
    const get = vi.fn(async (_path: string) => ({
      value: [
        { Id: 'c-1', DisplayName: 'Red category', Color: 'Preset0' },
        { Id: 'c-2', DisplayName: 'Follow up', Color: 'Preset4' },
      ],
    }));
    const client = { get } as unknown as OutlookClient;
    const h = await createTestHarness((s: McpServer) => registerDirectoryTools(s, client));
    const out = parseToolResult<{ count?: number; categories?: unknown[] }>(
      await h.callTool('outlook_list_categories', {}),
    );
    expect(get.mock.calls[0][0]).toBe('/me/outlook/masterCategories');
    expect(out.count).toBe(2);
    expect(out.categories).toEqual([
      { name: 'Red category', color: 'Preset0' },
      { name: 'Follow up', color: 'Preset4' },
    ]);
    await h.close();
  });
});
