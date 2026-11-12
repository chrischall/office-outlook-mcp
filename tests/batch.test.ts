import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerRuleTools } from '../src/tools/rules.js';
import { registerFolderTools } from '../src/tools/folders.js';
import { registerBatchTools } from '../src/tools/batch.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_batch` (mcp-utils registerBatchTool, docs/CONFIRM-TOKEN.md §14):
 * several mailbox writes under ONE preview and ONE approval, in order,
 * stopping at the first failure. The client is an in-memory mailbox, so a
 * later op sees what an earlier op created. Fixtures are synthetic.
 */

type Obj = Record<string, unknown>;
const RULES = '/me/MailFolders/inbox/MessageRules';
const ALIASES: Record<string, string> = { msgfolderroot: 'root-id', inbox: 'inbox-id', archive: 'archive-id' };

function stub(opts: { failRuleCreate?: boolean } = {}) {
  const folders: Obj[] = [
    { Id: 'root-id', DisplayName: 'Top', ParentFolderId: 'store' },
    { Id: 'inbox-id', DisplayName: 'Inbox', ParentFolderId: 'root-id' },
    { Id: 'archive-id', DisplayName: 'Archive', ParentFolderId: 'root-id' },
  ];
  const rules: Obj[] = [];
  const calls: { method: string; path: string; body?: unknown }[] = [];
  let n = 0;
  const resolve = (ref: string) => ALIASES[ref.toLowerCase()] ?? ref;
  const seg = (path: string, i: number) => decodeURIComponent((path.split('/')[i] ?? '').split('?')[0] ?? '');

  const get = vi.fn(async (path: string): Promise<unknown> => {
    calls.push({ method: 'GET', path });
    if (path === RULES) return { value: rules };
    if (path.startsWith(`${RULES}/`)) {
      const r = rules.find((x) => x.Id === seg(path, 5));
      if (!r) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      return r;
    }
    if (path.startsWith('/me/mailfolders/')) {
      const f = folders.find((x) => x.Id === resolve(seg(path, 3)));
      if (!f) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      return f;
    }
    if (path.startsWith('/me/messages/')) return { IsRead: true };
    throw new Error(`unexpected GET ${path}`);
  });

  const write = vi.fn(async (method: string, path: string, body?: unknown): Promise<unknown> => {
    calls.push({ method, path, body });
    const b = (body ?? {}) as Obj;
    if (method === 'POST' && path.endsWith('/childfolders')) {
      const parent = folders.find((x) => x.Id === resolve(seg(path, 3)));
      if (!parent) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      const made = { Id: `folder-${++n}`, DisplayName: b.DisplayName, ParentFolderId: parent.Id };
      folders.push(made);
      return made;
    }
    if (method === 'POST' && path === RULES) {
      if (opts.failRuleCreate) throw new ApiError(400, 'Outlook 400: ErrorInvalidRule');
      const made = { Id: `rule-${++n}`, ...b };
      rules.push(made);
      return made;
    }
    if (method === 'POST' && path.endsWith('/move')) return { Id: 'moved-1', ParentFolderId: b.DestinationId };
    if (method === 'PATCH') return { ...b };
    throw new Error(`unexpected ${method} ${path}`);
  });

  const client = { get, write } as unknown as OutlookClient;
  const writes = () => calls.filter((c) => c.method !== 'GET');
  return { client, calls, writes, folders, rules };
}

const harness = (client: OutlookClient) =>
  createTestHarness((s: McpServer) => {
    registerWriteTools(s, client);
    registerRuleTools(s, client);
    registerFolderTools(s, client);
    registerBatchTools(s, client);
  });

type PhaseOne = {
  status?: string;
  dispatched?: boolean;
  confirmToken?: string;
  confirmTool?: string;
  preview?: { action?: string; willSend?: Obj[] };
};
type BatchResult = {
  status: 'ok' | 'partial' | 'failed';
  results: { index: number; tool: string; applied: boolean; result?: Obj; error?: string }[];
  not_run?: number[];
  note?: string;
};

const ENV_KEYS = ['MCP_CONFIRM_MODE', 'MCP_CONFIRM_TTL_SECONDS', 'MCP_CONFIRM_SECRET', 'MCP_BATCH_MAX_OPS'] as const;
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

const FOLDER_THEN_RULE = [
  { tool: 'outlook_create_folder', args: { name: 'Receipts' } },
  {
    tool: 'outlook_create_rule',
    args: { name: 'File receipts', conditions: { from: ['shop@example.com'] }, actions: { markAsRead: true } },
    refs: { moveToFolder: '0.id' },
  },
];

async function confirmBatch(h: TestHarness, ops: unknown[]) {
  const first = parseToolResult<PhaseOne>(await h.callTool('outlook_batch', { ops }));
  return h.callTool('outlook_batch', { ops, confirmToken: first.confirmToken });
}

describe('outlook_batch', () => {
  it('previews a folder + a rule filing into it once, and writes nothing until confirmed', async () => {
    const s = stub();
    const h = await harness(s.client);
    const first = parseToolResult<PhaseOne>(await h.callTool('outlook_batch', { ops: FOLDER_THEN_RULE }));
    expect(first.status).toBe('confirmation-required');
    expect(first.dispatched).toBe(false);
    expect(first.confirmToken).toEqual(expect.any(String));
    const shown = first.preview?.willSend ?? [];
    expect(shown).toHaveLength(2);
    expect(shown[0]).toMatchObject({ tool: 'outlook_create_folder', method: 'POST', willSend: { DisplayName: 'Receipts' } });
    // The rule's folder does not exist yet: it is shown by its arguments, the ref named.
    expect(shown[1]).toMatchObject({ tool: 'outlook_create_rule', args: { moveToFolder: '<id from op 0>' } });
    expect(s.writes()).toHaveLength(0);
    await h.close();
  });

  it('runs the ops in order with the new folder id fed into the rule', async () => {
    const s = stub();
    const h = await harness(s.client);
    const res = await confirmBatch(h, FOLDER_THEN_RULE);
    expect(res.isError).toBeFalsy();
    const out = parseToolResult<BatchResult>(res);
    expect(out.status).toBe('ok');
    expect(out.results.map((r) => [r.tool, r.applied])).toEqual([
      ['outlook_create_folder', true],
      ['outlook_create_rule', true],
    ]);
    expect(s.writes().map((w) => [w.method, w.path])).toEqual([
      ['POST', '/me/mailfolders/msgfolderroot/childfolders'],
      ['POST', RULES],
    ]);
    expect(s.writes()[1].body).toMatchObject({
      DisplayName: 'File receipts',
      Actions: { MoveToFolder: 'folder-1', MarkAsRead: true },
    });
    expect(out.results[0].result).toMatchObject({ created: true, Id: 'folder-1' });
    await h.close();
  });

  it('stops at a failing op: partial, no rollback, and the rest never run', async () => {
    const s = stub({ failRuleCreate: true });
    const h = await harness(s.client);
    const ops = [...FOLDER_THEN_RULE, { tool: 'outlook_move_message', args: { id: 'm1', destination: 'archive' } }];
    const res = await confirmBatch(h, ops);
    expect(res.isError).toBe(true);
    const out = parseToolResult<BatchResult>(res);
    expect(out.status).toBe('partial');
    expect(out.results.map((r) => r.applied)).toEqual([true, false]);
    expect(out.results[1].error).toMatch(/400/);
    expect(out.not_run).toEqual([2]);
    expect(out.note).toMatch(/No rollback/);
    // The folder was made, the rule was attempted, the move never sent.
    expect(s.writes().map((w) => w.path)).toEqual(['/me/mailfolders/msgfolderroot/childfolders', RULES]);
    await h.close();
  });

  it('refuses a send: only mailbox writes are batchable', async () => {
    const s = stub();
    const h = await harness(s.client);
    const res = await h.callTool('outlook_batch', {
      ops: [{ tool: 'outlook_send_message', args: { to: ['a@example.com'], subject: 's', body: 'b' } }],
    });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/not a write this batch can run/);
    expect(s.writes()).toHaveLength(0);
    await h.close();
  });

  it('caps a batch at 100 ops', async () => {
    const s = stub();
    const h = await harness(s.client);
    const ops = Array.from({ length: 101 }, (_, i) => ({ tool: 'outlook_mark_read', args: { id: `m${i}`, isRead: true } }));
    const res = await h.callTool('outlook_batch', { ops });
    expect(res.isError).toBe(true);
    expect(s.writes()).toHaveLength(0);
    await h.close();
  });

  it('can be confirmed with the token alone through outlook_confirm', async () => {
    const s = stub();
    const h = await harness(s.client);
    const first = parseToolResult<PhaseOne>(await h.callTool('outlook_batch', { ops: FOLDER_THEN_RULE }));
    expect(first.confirmTool).toBe('outlook_confirm');
    const res = await h.callTool('outlook_confirm', { confirmToken: first.confirmToken });
    expect(res.isError).toBeFalsy();
    expect(parseToolResult<BatchResult>(res).status).toBe('ok');
    expect(s.writes()).toHaveLength(2);
    await h.close();
  });

  it('lists exactly the mailbox writes, and points bulk moves at outlook_move_messages', async () => {
    const h = await harness(stub().client);
    const { tools } = await h.client.listTools();
    const batch = tools.find((t) => t.name === 'outlook_batch');
    expect(batch).toBeDefined();
    const d = batch!.description ?? '';
    expect(d).toMatch(/outlook_move_messages/);
    expect(d).toMatch(/no rollback/i);
    for (const name of [
      'outlook_move_message',
      'outlook_mark_read',
      'outlook_update_message',
      'outlook_create_rule',
      'outlook_update_rule',
      'outlook_delete_rule',
      'outlook_create_folder (produces id)',
      'outlook_update_folder',
    ]) {
      expect(d, name).toContain(name);
    }
    for (const name of ['outlook_send_message', 'outlook_reply', 'outlook_create_event', 'outlook_delete_folder', 'outlook_set_auto_reply']) {
      expect(d, name).not.toContain(`${name},`);
      expect(d, name).not.toMatch(new RegExp(`${name}\\.`));
    }
    await h.close();
  });
});

describe('the converted writes stay single tools', () => {
  it('outlook_create_folder still creates without asking (nothing reaches anyone else)', async () => {
    const s = stub();
    const h = await harness(s.client);
    const out = parseToolResult<Obj>(await h.callTool('outlook_create_folder', { name: 'Receipts', parent: 'inbox' }));
    expect(out).toMatchObject({ created: true, Id: 'folder-1', Name: 'Receipts', Parent: 'inbox-id' });
    expect(s.writes()).toHaveLength(1);
    await h.close();
  });

  it('outlook_create_rule takes a top-level moveToFolder, the same as actions.moveToFolder', async () => {
    const s = stub();
    const h = await harness(s.client);
    const args = { name: 'r', conditions: { from: ['a@example.com'] }, actions: {}, moveToFolder: 'archive' };
    const first = parseToolResult<PhaseOne>(await h.callTool('outlook_create_rule', args));
    expect(first.status).toBe('confirmation-required');
    await h.callTool('outlook_create_rule', { ...args, confirmToken: first.confirmToken });
    expect(s.writes()[0].body).toMatchObject({ Actions: { MoveToFolder: 'archive-id' } });
    await h.close();
  });

  it('outlook_create_rule refuses moveToFolder given twice', async () => {
    const h = await harness(stub().client);
    const res = await h.callTool('outlook_create_rule', {
      name: 'r',
      conditions: { from: ['a@example.com'] },
      actions: { moveToFolder: 'inbox' },
      moveToFolder: 'archive',
    });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/moveToFolder/);
    await h.close();
  });
});
