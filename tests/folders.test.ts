import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerFolderTools } from '../src/tools/folders.js';
import type { OutlookClient } from '../src/client.js';

/**
 * Mail folder management. Paths, status codes and field names match what the
 * v2.0 MailFolder endpoints did live on 2026-11-10: a duplicate name answers 409
 * ErrorFolderExists, a malformed parent 400, a move keeps the folder's Id, and a
 * delete removes the folder outright (re-read 404s; nothing in Deleted Items).
 * Fixtures are synthetic.
 */

type Obj = Record<string, unknown>;
type Folder = Obj & { Id: string; DisplayName: string; ParentFolderId: string };

const ROOT = 'root-id';
const DELETED = 'deleted-id';
const RULES = '/me/MailFolders/inbox/MessageRules';

function folder(over: Partial<Folder> & { Id: string }): Folder {
  return {
    '@odata.id': 'x',
    DisplayName: over.Id,
    ParentFolderId: ROOT,
    ChildFolderCount: 0,
    UnreadItemCount: 0,
    TotalItemCount: 0,
    ...over,
  };
}

const ALIASES: Record<string, string> = { msgfolderroot: ROOT, inbox: 'inbox-id', deleteditems: DELETED, archive: 'archive-id' };

/**
 * A mailbox held in memory, so every read-back sees what the write did.
 */
function stub(
  opts: {
    folders?: Folder[];
    rules?: Obj[];
    persist?: boolean;
    writeError?: Error;
  } = {},
) {
  let folders: Folder[] = (
    opts.folders ?? [
      folder({ Id: ROOT, DisplayName: 'Top of Information Store', ParentFolderId: 'store' }),
      folder({ Id: 'inbox-id', DisplayName: 'Inbox' }),
      folder({ Id: DELETED, DisplayName: 'Deleted Items' }),
      folder({ Id: 'archive-id', DisplayName: 'Archive' }),
      folder({ Id: 'f1', DisplayName: 'Alerts', ParentFolderId: 'inbox-id', TotalItemCount: 42, UnreadItemCount: 7, ChildFolderCount: 1 }),
      folder({ Id: 'f1a', DisplayName: 'Old alerts', ParentFolderId: 'f1' }),
    ]
  ).map((f) => ({ ...f }));
  const persist = opts.persist !== false;
  const resolve = (ref: string) => ALIASES[ref.toLowerCase()] ?? ref;
  const find = (ref: string) => {
    const f = folders.find((x) => x.Id === resolve(ref));
    if (!f) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
    return f;
  };
  const seg = (path: string, i: number) => decodeURIComponent((path.split('/')[i] ?? '').split('?')[0] ?? '');

  const get = vi.fn(async (path: string): Promise<unknown> => {
    if (path === RULES) return { value: opts.rules ?? [] };
    if (path.startsWith('/me/mailfolders/')) return find(seg(path, 3));
    throw new Error(`unexpected GET ${path}`);
  });
  const write = vi.fn(async (method: string, path: string, body?: Obj) => {
    if (opts.writeError) throw opts.writeError;
    const id = seg(path, 3);
    const tail = path.split('/')[4];
    if (method === 'POST' && tail === 'childfolders') {
      if (body?.DisplayName === 'bad\\name') throw new ApiError(400, 'Outlook 400: ErrorInvalidRequest');
      const parent = find(id);
      if (folders.some((f) => f.ParentFolderId === parent.Id && f.DisplayName === body?.DisplayName)) {
        throw new ApiError(409, 'Outlook 409: ErrorFolderExists');
      }
      const made = folder({ Id: 'new-folder', DisplayName: String(body?.DisplayName), ParentFolderId: parent.Id });
      if (persist) folders.push(made);
      return made;
    }
    const target = find(id);
    if (method === 'POST' && tail === 'move') {
      const dest = find(String(body?.DestinationId));
      const moved = { ...target, ParentFolderId: dest.Id };
      if (persist) folders = folders.map((f) => (f.Id === target.Id ? moved : f));
      return moved;
    }
    if (method === 'PATCH') {
      if (persist) folders = folders.map((f) => (f.Id === target.Id ? { ...f, ...body } : f));
      return { ...target, ...body };
    }
    if (method === 'DELETE') {
      if (!persist) return undefined;
      folders = folders.filter((f) => f.Id !== target.Id);
      return undefined;
    }
    throw new Error(`unexpected ${method} ${path}`);
  });
  return { client: { get, write } as unknown as OutlookClient, write, get, now: () => folders };
}

const harness = (client: OutlookClient) => createTestHarness((s: McpServer) => registerFolderTools(s, client));

type Result = Obj & { status?: string; confirmToken?: string; warning?: string };

async function call(h: TestHarness, tool: string, args: Obj): Promise<Result> {
  return parseToolResult<Result>(await h.callTool(tool, args));
}

async function confirmed(h: TestHarness, tool: string, args: Obj): Promise<Result> {
  const first = await call(h, tool, args);
  expect(first.status).toBe('confirmation-required');
  return call(h, tool, { ...args, confirmToken: first.confirmToken });
}

describe('outlook_create_folder', () => {
  it('creates a top-level folder under the mailbox root, with no confirmation', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const res = await call(h, 'outlook_create_folder', { name: 'Alerts 2' });
    expect(write.mock.calls[0]).toEqual(['POST', '/me/mailfolders/msgfolderroot/childfolders', { DisplayName: 'Alerts 2' }]);
    expect(res).toMatchObject({ created: true, Id: 'new-folder', Name: 'Alerts 2' });
    expect(res).not.toHaveProperty('@odata.id');
    expect(now().some((f) => f.DisplayName === 'Alerts 2' && f.ParentFolderId === ROOT)).toBe(true);
    await h.close();
  });

  it('creates a subfolder under a parent given by id or well-known name', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    await call(h, 'outlook_create_folder', { name: 'Builds', parent: 'inbox' });
    expect(write.mock.calls[0]?.slice(0, 2)).toEqual(['POST', '/me/mailfolders/inbox/childfolders']);
    await call(h, 'outlook_create_folder', { name: 'Nightly', parent: 'f1' });
    expect(write.mock.calls[1]?.slice(0, 2)).toEqual(['POST', '/me/mailfolders/f1/childfolders']);
    await h.close();
  });

  it('explains a name that is already taken in that parent', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_create_folder', { name: 'Alerts', parent: 'inbox' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/already exists/i);
    await h.close();
  });

  it('does not blame a missing parent when none was given', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_create_folder', { name: 'bad\\name' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).not.toMatch(/undefined|No mail folder matches/);
    await h.close();
  });

  it('explains a parent that does not exist', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_create_folder', { name: 'x', parent: 'nope' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/outlook_list_folders/);
    await h.close();
  });

  it('reports an unknown outcome instead of inviting a duplicate', async () => {
    const { client } = stub({ writeError: new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await call(h, 'outlook_create_folder', { name: 'x' });
    expect(res.status).toBe('unknown');
    expect(res.warning).toMatch(/outlook_list_folders/);
    await h.close();
  });
});

describe('outlook_update_folder', () => {
  it('renames after confirmation, showing the current folder in the preview', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const first = await call(h, 'outlook_update_folder', { id: 'f1', name: 'Pager' });
    expect(JSON.stringify(first)).toMatch(/Alerts/);
    expect(write).not.toHaveBeenCalled();
    const res = await call(h, 'outlook_update_folder', { id: 'f1', name: 'Pager', confirmToken: first.confirmToken });
    expect(write.mock.calls).toEqual([['PATCH', '/me/mailfolders/f1', { DisplayName: 'Pager' }]]);
    expect(res).toMatchObject({ updated: true, Id: 'f1', Name: 'Pager' });
    expect(now().find((f) => f.Id === 'f1')?.DisplayName).toBe('Pager');
    await h.close();
  });

  it('moves under a new parent, resolving a well-known name', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_folder', { id: 'f1', parent: 'archive' });
    expect(write.mock.calls).toEqual([['POST', '/me/mailfolders/f1/move', { DestinationId: 'archive-id' }]]);
    expect(res).toMatchObject({ updated: true, Id: 'f1' });
    expect(now().find((f) => f.Id === 'f1')?.ParentFolderId).toBe('archive-id');
    await h.close();
  });

  it('moves to the top level with parent "msgfolderroot"', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const first = await call(h, 'outlook_update_folder', { id: 'f1', parent: 'msgfolderroot' });
    expect(JSON.stringify(first)).toMatch(/top level/);
    await call(h, 'outlook_update_folder', { id: 'f1', parent: 'msgfolderroot', confirmToken: first.confirmToken });
    expect(write.mock.calls[0]).toEqual(['POST', '/me/mailfolders/f1/move', { DestinationId: ROOT }]);
    await h.close();
  });

  it('renames and moves in one call', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_folder', { id: 'f1', name: 'Pager', parent: 'archive' });
    expect(write.mock.calls.map((c) => c.slice(0, 2))).toEqual([
      ['PATCH', '/me/mailfolders/f1'],
      ['POST', '/me/mailfolders/f1/move'],
    ]);
    expect(res).toMatchObject({ updated: true, Name: 'Pager' });
    expect(now().find((f) => f.Id === 'f1')).toMatchObject({ DisplayName: 'Pager', ParentFolderId: 'archive-id' });
    await h.close();
  });

  it('previews the requests it will actually send, in order', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await call(h, 'outlook_update_folder', { id: 'f1', name: 'Pager', parent: 'archive' });
    expect((res.preview as Obj).willSend).toEqual([
      { method: 'PATCH', path: '/me/mailfolders/f1', body: { DisplayName: 'Pager' } },
      { method: 'POST', path: '/me/mailfolders/f1/move', body: { DestinationId: 'archive-id' } },
    ]);
    await h.close();
  });

  it('returns the id Outlook gives the moved folder', async () => {
    const { client, write } = stub();
    write.mockImplementationOnce(async () => folder({ Id: 'f1-moved', DisplayName: 'Alerts', ParentFolderId: 'archive-id' }));
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_folder', { id: 'f1', parent: 'archive' });
    expect(res).toMatchObject({ Id: 'f1-moved' });
    await h.close();
  });

  it.each([
    ['nothing to change', { id: 'f1' }],
    ['a well-known folder', { id: 'Inbox', name: 'Mail' }],
    ['a move into itself', { id: 'f1', parent: 'f1' }],
  ])('refuses %s before asking anything', async (_label, args) => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_folder', args);
    expect(res.isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('warns when Outlook accepts the rename but the read-back disagrees', async () => {
    const { client } = stub({ persist: false });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_folder', { id: 'f1', name: 'Pager' });
    expect(res).toMatchObject({ updated: false });
    expect(res.warning).toBeTruthy();
    await h.close();
  });
});

describe('outlook_delete_folder', () => {
  it('shows what goes with the folder, then deletes and verifies with a 404', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const first = await call(h, 'outlook_delete_folder', { id: 'f1' });
    const text = JSON.stringify(first);
    expect(text).toMatch(/Alerts/);
    expect(text).toMatch(/42/);
    expect(text).toMatch(/subfolder/i);
    expect(text).toMatch(/does not go to Deleted Items/);
    expect(write).not.toHaveBeenCalled();
    const res = await call(h, 'outlook_delete_folder', { id: 'f1', confirmToken: first.confirmToken });
    expect(write.mock.calls).toEqual([['DELETE', '/me/mailfolders/f1']]);
    expect(res).toMatchObject({ deleted: true, Id: 'f1' });
    expect(now().some((f) => f.Id === 'f1')).toBe(false);
    await h.close();
  });

  it('names the inbox rules that file mail into it', async () => {
    const { client } = stub({
      rules: [
        { Id: 'r1', DisplayName: 'Pager duty', Actions: { MoveToFolder: 'f1' } },
        { Id: 'r2', DisplayName: 'Copy alerts', Actions: { CopyToFolder: 'f1' } },
        { Id: 'r3', DisplayName: 'Unrelated', Actions: { MoveToFolder: 'archive-id' } },
      ],
    });
    const h = await harness(client);
    const text = JSON.stringify(await call(h, 'outlook_delete_folder', { id: 'f1' }));
    expect(text).toMatch(/Pager duty/);
    expect(text).toMatch(/Copy alerts/);
    expect(text).not.toMatch(/Unrelated/);
    await h.close();
  });

  it('still previews when the rules cannot be read', async () => {
    const { client, get } = stub();
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    get.mockImplementation(async (p: string) => {
      if (p === RULES) throw new ApiError(503, 'Outlook 503');
      return real(p);
    });
    const h = await harness(client);
    const res = await call(h, 'outlook_delete_folder', { id: 'f1' });
    expect(res.status).toBe('confirmation-required');
    await h.close();
  });

  it('refuses a well-known folder before asking anything', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_delete_folder', { id: 'deleteditems' });
    expect(res.isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('does not claim success when the folder is still where it was', async () => {
    const { client } = stub({ persist: false });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_delete_folder', { id: 'f1' });
    expect(res).toMatchObject({ deleted: false });
    expect(res.warning).toBeTruthy();
    await h.close();
  });
});

describe('deleting a large folder', () => {
  it('says the delete may still be running when Outlook does not answer in time, and not to resend', async () => {
    const { client } = stub({ writeError: new WriteOutcomeUnknownError('Outlook', 'DELETE', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_delete_folder', { id: 'f1' });
    expect(res.status).toBe('unknown');
    expect(res.warning).toMatch(/still running|keeps running/i);
    expect(res.warning).toMatch(/Total/);
    expect(res.warning).toMatch(/not resend|do not retry|don't resend/i);
    await h.close();
  });
});

describe('a folder write that lands but cannot be re-read', () => {
  it.each([
    ['outlook_update_folder', { id: 'f1', name: 'Pager' }],
    ['outlook_delete_folder', { id: 'f1' }],
  ])('%s says to check before retrying, and does not error', async (tool, args) => {
    const { client, write, get } = stub();
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    get.mockImplementation(async (p: string) => {
      if (write.mock.calls.length > 0) throw new ApiError(503, 'Outlook 503');
      return real(p);
    });
    const h = await harness(client);
    const res = await confirmed(h, tool, args);
    expect(write).toHaveBeenCalledTimes(1);
    expect(res.warning).toMatch(/before retrying/i);
    await h.close();
  });
});
