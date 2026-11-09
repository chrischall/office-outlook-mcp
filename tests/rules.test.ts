import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerRuleTools } from '../src/tools/rules.js';
import type { OutlookClient } from '../src/client.js';

/**
 * Inbox rules and automatic replies. Field names match what
 * /me/MailFolders/inbox/MessageRules and /me/MailboxSettings returned live on
 * 2026-11-09. Fixtures are synthetic (example.test addresses).
 */

type Obj = Record<string, unknown>;
const RULES = '/me/MailFolders/inbox/MessageRules';

function rule(over: Obj = {}): Obj {
  return {
    '@odata.id': 'x',
    Id: 'r1',
    DisplayName: 'Newsletters',
    Sequence: 3,
    IsEnabled: true,
    HasError: false,
    IsReadOnly: false,
    Conditions: { FromAddresses: [{ EmailAddress: { Address: 'news@example.test' } }] },
    Actions: { MoveToFolder: 'folder-id', StopProcessingRules: true },
    ...over,
  };
}

function reply(over: Obj = {}): Obj {
  return {
    Status: 'Disabled',
    ExternalAudience: 'None',
    InternalReplyMessage: 'old internal text',
    ExternalReplyMessage: 'old external text',
    ScheduledStartDateTime: { DateTime: '2026-01-01T00:00:00.0000000', TimeZone: 'UTC' },
    ScheduledEndDateTime: { DateTime: '2026-01-02T00:00:00.0000000', TimeZone: 'UTC' },
    ...over,
  };
}

/** A mailbox held in memory, so every read-back sees what the write did. */
function stub(
  rules: Obj[] = [rule()],
  opts: { auto?: Obj; persist?: boolean; writeError?: Error } = {},
) {
  let stored = rules.map((r) => ({ ...r }));
  let auto = opts.auto ?? reply();
  const persist = opts.persist !== false;
  const get = vi.fn(async (path: string) => {
    if (path === '/me/MailboxSettings') return { TimeZone: 'Eastern Standard Time', AutomaticRepliesSetting: auto };
    if (path === RULES) return { value: stored };
    if (path.startsWith('/me/mailfolders/')) {
      const name = decodeURIComponent((path.split('/')[3] ?? '').split('?')[0] ?? '');
      if (name === 'missing') throw new ApiError(404, 'Outlook 404: ErrorInvalidIdMalformed');
      return { Id: `id-of-${name}`, DisplayName: name };
    }
    if (path.startsWith(`${RULES}/`)) {
      const id = decodeURIComponent(path.slice(RULES.length + 1));
      const found = stored.find((r) => r.Id === id);
      if (!found) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      return found;
    }
    throw new Error(`unexpected GET ${path}`);
  });
  const write = vi.fn(async (method: string, path: string, body?: Obj) => {
    if (opts.writeError) throw opts.writeError;
    if (!persist) return method === 'POST' ? { Id: 'new-rule', ...body } : undefined;
    if (path === '/me/MailboxSettings') {
      auto = body?.AutomaticRepliesSetting as Obj;
      return undefined;
    }
    if (method === 'POST') {
      const created = { Id: 'new-rule', HasError: false, IsReadOnly: false, ...body };
      stored.push(created);
      return created;
    }
    const id = decodeURIComponent(path.slice(RULES.length + 1));
    if (method === 'DELETE') stored = stored.filter((r) => r.Id !== id);
    else stored = stored.map((r) => (r.Id === id ? { ...r, ...body } : r));
    return undefined;
  });
  return { client: { get, write } as unknown as OutlookClient, write, rulesNow: () => stored };
}

const harness = (client: OutlookClient) =>
  createTestHarness((s: McpServer) => registerRuleTools(s, client));

type Result = Obj & { status?: string; confirmToken?: string; warning?: string };

async function preview(h: TestHarness, tool: string, args: Obj): Promise<Result> {
  return parseToolResult<Result>(await h.callTool(tool, args));
}

async function confirmed(h: TestHarness, tool: string, args: Obj): Promise<Result> {
  const first = await preview(h, tool, args);
  return parseToolResult<Result>(await h.callTool(tool, { ...args, confirmToken: first.confirmToken }));
}

const createArgs = { name: 'Receipts', conditions: { subjectContains: ['receipt'] }, actions: { markAsRead: true } };

describe('outlook_list_rules', () => {
  it('returns the rules in the order Outlook applies them, without OData noise', async () => {
    const { client } = stub([rule({ Id: 'b', Sequence: 2 }), rule({ Id: 'a', Sequence: 1 })]);
    const h = await harness(client);
    const out = parseToolResult<{ items: Obj[] }>(await h.callTool('outlook_list_rules', {}));
    expect(out.items.map((r) => r.Id)).toEqual(['a', 'b']);
    expect(out.items[0]).not.toHaveProperty('@odata.id');
    expect(out.items[0]).toMatchObject({ DisplayName: 'Newsletters', IsEnabled: true, Actions: { StopProcessingRules: true } });
    await h.close();
  });
});

describe('outlook_create_rule', () => {
  it('previews and writes nothing without a token', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await preview(h, 'outlook_create_rule', createArgs);
    expect(res.status).toBe('confirmation-required');
    expect(JSON.stringify(res)).toMatch(/Receipts/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('maps the arguments to Outlook fields, sequences it last and verifies by re-reading', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_create_rule', {
      name: 'Receipts',
      conditions: { from: ['a@example.test'], subjectContains: ['receipt'], hasAttachments: true },
      exceptions: { senderContains: ['noreply'] },
      actions: { moveToFolder: 'archive', markAsRead: true, assignCategories: ['Money'], stopProcessingRules: true },
    });
    expect(write).toHaveBeenCalledTimes(1);
    expect(write.mock.calls[0]?.slice(0, 2)).toEqual(['POST', RULES]);
    expect(write.mock.calls[0]?.[2]).toEqual({
      DisplayName: 'Receipts',
      Sequence: 4,
      IsEnabled: true,
      Conditions: {
        FromAddresses: [{ EmailAddress: { Address: 'a@example.test' } }],
        SubjectContains: ['receipt'],
        HasAttachments: true,
      },
      Exceptions: { SenderContains: ['noreply'] },
      Actions: { MoveToFolder: 'id-of-archive', MarkAsRead: true, AssignCategories: ['Money'], StopProcessingRules: true },
    });
    expect(res).toMatchObject({ created: true, Id: 'new-rule' });
    await h.close();
  });

  it('says plainly when a rule sends mail to someone else, naming them', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await preview(h, 'outlook_create_rule', {
      name: 'Fwd',
      conditions: { subjectContains: ['invoice'] },
      actions: { forwardTo: ['out@example.test'], redirectTo: ['other@example.test'] },
    });
    const text = JSON.stringify(res);
    expect(text).toMatch(/out@example\.test/);
    expect(text).toMatch(/other@example\.test/);
    expect(text).toMatch(/every matching message/i);
    await h.close();
  });

  it.each([
    ['no conditions', { name: 'x', conditions: {}, actions: { markAsRead: true } }],
    ['no actions', { name: 'x', conditions: { subjectContains: ['a'] }, actions: {} }],
  ])('refuses %s before asking anything', async (_label, args) => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_create_rule', args);
    expect(res.isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('refuses a folder that does not resolve before asking anything', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_create_rule', {
      name: 'x',
      conditions: { subjectContains: ['a'] },
      actions: { moveToFolder: 'missing' },
    });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/outlook_list_folders/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('reports an unknown outcome instead of inviting a duplicate', async () => {
    const { client } = stub([rule()], { writeError: new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_create_rule', createArgs);
    expect(res.status).toBe('unknown');
    await h.close();
  });
});

describe('a write that lands but cannot be re-read', () => {
  const flaky = () => {
    const { client, write } = stub();
    const get = client.get as ReturnType<typeof vi.fn>;
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    get.mockImplementation(async (p: string) => {
      if (write.mock.calls.length > 0) throw new ApiError(503, 'Outlook 503');
      return real(p);
    });
    return { client, write };
  };

  it.each([
    ['outlook_create_rule', createArgs],
    ['outlook_update_rule', { id: 'r1', enabled: false }],
    ['outlook_delete_rule', { id: 'r1' }],
    ['outlook_set_auto_reply', { status: 'alwaysEnabled' }],
  ])('%s says to check before retrying, and does not error', async (tool, args) => {
    const { client, write } = flaky();
    const h = await harness(client);
    const res = await confirmed(h, tool, args);
    expect(write).toHaveBeenCalledTimes(1);
    expect(res.warning).toMatch(/before retrying/i);
    await h.close();
  });
});

describe('outlook_update_rule', () => {
  it('disables a rule, sending only that field', async () => {
    const { client, write, rulesNow } = stub();
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_rule', { id: 'r1', enabled: false });
    expect(write.mock.calls[0]).toEqual(['PATCH', `${RULES}/r1`, { IsEnabled: false }]);
    expect(rulesNow()[0]?.IsEnabled).toBe(false);
    expect(res).toMatchObject({ updated: true });
    await h.close();
  });

  it('warns when Outlook accepts the change but the read-back disagrees', async () => {
    const { client } = stub([rule()], { persist: false });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_update_rule', { id: 'r1', enabled: false });
    expect(res).toMatchObject({ updated: false });
    expect(res.warning).toBeTruthy();
    await h.close();
  });

  it('shows the current rule in the preview, and that actions are replaced whole', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await preview(h, 'outlook_update_rule', { id: 'r1', actions: { markAsRead: true } });
    const text = JSON.stringify(res);
    expect(text).toMatch(/Newsletters/);
    expect(text).toMatch(/replace/i);
    await h.close();
  });

  it('refuses a read-only rule and an empty change', async () => {
    const { client, write } = stub([rule({ IsReadOnly: true })]);
    const h = await harness(client);
    expect((await h.callTool('outlook_update_rule', { id: 'r1', enabled: false })).isError).toBe(true);
    expect((await h.callTool('outlook_update_rule', { id: 'r1' })).isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });
});

describe('outlook_delete_rule', () => {
  it('names the rule in the preview, then deletes and verifies with a 404', async () => {
    const { client, write, rulesNow } = stub();
    const h = await harness(client);
    expect(JSON.stringify(await preview(h, 'outlook_delete_rule', { id: 'r1' }))).toMatch(/Newsletters/);
    expect(write).not.toHaveBeenCalled();
    const res = await confirmed(h, 'outlook_delete_rule', { id: 'r1' });
    expect(res).toMatchObject({ deleted: true });
    expect(rulesNow()).toHaveLength(0);
    await h.close();
  });

  it('does not claim success when the rule is still there', async () => {
    const { client } = stub([rule()], { persist: false });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_delete_rule', { id: 'r1' });
    expect(res).toMatchObject({ deleted: false });
    expect(res.warning).toBeTruthy();
    await h.close();
  });
});

describe('outlook_set_auto_reply', () => {
  it('shows the stored text it would send when turning on with no new message', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const text = JSON.stringify(await preview(h, 'outlook_set_auto_reply', { status: 'alwaysEnabled' }));
    expect(text).toMatch(/old internal text/);
    expect(text).toMatch(/old external text/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('merges over the stored setting, so a partial change keeps the rest', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_set_auto_reply', {
      status: 'alwaysEnabled',
      internalMessage: 'Away',
      externalAudience: 'all',
    });
    expect(write.mock.calls[0]).toEqual([
      'PATCH',
      '/me/MailboxSettings',
      {
        AutomaticRepliesSetting: reply({
          Status: 'AlwaysEnabled',
          ExternalAudience: 'All',
          InternalReplyMessage: 'Away',
        }),
      },
    ]);
    expect(res).toMatchObject({ updated: true, AutomaticRepliesSetting: { Status: 'AlwaysEnabled' } });
    await h.close();
  });

  it('schedules in the mailbox zone by default, and in UTC for a value with an offset', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    await confirmed(h, 'outlook_set_auto_reply', {
      status: 'scheduled',
      start: '2026-12-20T09:00:00',
      end: '2026-12-27T17:00:00-05:00',
    });
    const sent = (write.mock.calls[0]?.[2] as Obj).AutomaticRepliesSetting as Obj;
    expect(sent.Status).toBe('Scheduled');
    expect(sent.ScheduledStartDateTime).toEqual({ DateTime: '2026-12-20T09:00:00', TimeZone: 'Eastern Standard Time' });
    expect(sent.ScheduledEndDateTime).toEqual({ DateTime: '2026-12-27T22:00:00', TimeZone: 'UTC' });
    await h.close();
  });

  it.each([
    ['scheduled with no dates', { status: 'scheduled' }],
    ['an end before the start', { status: 'scheduled', start: '2026-12-27T09:00:00', end: '2026-12-20T09:00:00' }],
    ['nothing to change', {}],
  ])('refuses %s', async (_label, args) => {
    const { client, write } = stub();
    const h = await harness(client);
    expect((await h.callTool('outlook_set_auto_reply', args)).isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('warns when Outlook accepts the change but the status did not move', async () => {
    const { client } = stub([], { persist: false });
    const h = await harness(client);
    const res = await confirmed(h, 'outlook_set_auto_reply', { status: 'alwaysEnabled' });
    expect(res).toMatchObject({ updated: false });
    expect(res.warning).toBeTruthy();
    await h.close();
  });
});
