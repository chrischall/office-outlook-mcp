import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import type { OutlookClient } from '../src/client.js';

/**
 * The confirm-token flow. A harness created WITHOUT an elicitation handler is a
 * client that cannot be prompted (claude.ai, Claude Desktop): under the default
 * `MCP_CONFIRM_MODE=ask-user` the first call writes nothing and returns a
 * preview plus a `confirmToken`, and only a repeat call carrying that token
 * performs the write.
 */

function stubClient() {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  const client = {
    get: vi.fn(async (path: string) => {
      calls.push({ method: 'GET', path });
      if (path.includes('MailboxSettings')) return { TimeZone: 'Eastern Standard Time' };
      if (path.includes('$select=IsRead')) return { IsRead: true };
      return { value: [] };
    }),
    write: vi.fn(async (method: string, path: string, body?: unknown) => {
      calls.push({ method, path, body });
      return { Id: 'new-1', WebLink: 'w' };
    }),
  } as unknown as OutlookClient;
  const writes = () => calls.filter((c) => c.method !== 'GET');
  return { client, calls, writes };
}

type PhaseOne = {
  status?: string;
  confirmed?: boolean;
  dispatched?: boolean;
  action?: string;
  confirmToken?: string;
  preview?: {
    action?: string;
    method?: string;
    path?: string;
    willSend?: unknown;
  };
};

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

const gated = [
  [
    'outlook_send_mail',
    { to: ['a@example.com'], cc: ['c@example.com'], subject: 's', body: 'b' },
    'POST',
    '/me/sendmail',
  ],
  ['outlook_create_draft', { to: ['a@example.com'], subject: 's', body: 'b' }, 'POST', '/me/messages'],
  ['outlook_mark_read', { id: 'm1', isRead: true }, 'PATCH', '/me/messages/m1'],
  ['outlook_move_message', { id: 'm1', destination: 'archive' }, 'POST', '/me/messages/m1/move'],
  [
    'outlook_create_event',
    { subject: 's', start: '2026-09-22T15:00:00', end: '2026-09-22T16:00:00' },
    'POST',
    '/me/events',
  ],
] as const;

describe('every write tool is gated by a confirm token', () => {
  it.each(gated)('%s: phase 1 previews and writes nothing; phase 2 writes once', async (name, args, method, path) => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));

    const first = parseToolResult<PhaseOne>(await h.callTool(name, { ...args }));
    expect(first.status).toBe('confirmation-required');
    expect(first.dispatched).toBe(false);
    expect(first.confirmToken).toEqual(expect.any(String));
    // The preview carries everything the old dry run did.
    expect(first.preview?.method).toBe(method);
    expect(first.preview?.path).toBe(path);
    expect(first.preview?.action).toEqual(expect.any(String));
    expect(first.preview?.willSend).toBeDefined();
    expect(writes()).toHaveLength(0);

    const second = await h.callTool(name, { ...args, confirmToken: first.confirmToken });
    expect(second.isError).toBeFalsy();
    expect(writes()).toHaveLength(1);
    expect(writes()[0]).toMatchObject({ method, path, body: first.preview?.willSend });
    await h.close();
  });

  it('exposes confirmToken and no confirm parameter', async () => {
    const { client } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const { tools } = await h.client.listTools();
    expect(tools).toHaveLength(5);
    for (const tool of tools) {
      const props = tool.inputSchema.properties ?? {};
      expect(props).toHaveProperty('confirmToken');
      expect(props).not.toHaveProperty('confirm');
      expect(tool.description).toMatch(/confirmToken/);
      expect(tool.description).not.toMatch(/confirm:\s*true/);
    }
    await h.close();
  });

  it('refuses a replayed token and writes nothing more', async () => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const args = { to: ['a@example.com'], subject: 's', body: 'b' };
    const first = parseToolResult<PhaseOne>(await h.callTool('outlook_send_mail', args));
    await h.callTool('outlook_send_mail', { ...args, confirmToken: first.confirmToken });
    expect(writes()).toHaveLength(1);

    const replay = await h.callTool('outlook_send_mail', { ...args, confirmToken: first.confirmToken });
    expect(replay.isError).toBe(true);
    expect(parseToolResult<{ error?: string }>(replay).error).toBe('TOKEN_REUSED');
    expect(writes()).toHaveLength(1);
    await h.close();
  });

  it('refuses a token when an argument changed between the phases', async () => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const first = parseToolResult<PhaseOne>(
      await h.callTool('outlook_send_mail', { to: ['a@example.com'], subject: 's', body: 'b' }),
    );
    const changed = await h.callTool('outlook_send_mail', {
      to: ['x@evil.test'],
      subject: 's',
      body: 'b',
      confirmToken: first.confirmToken,
    });
    expect(changed.isError).toBe(true);
    expect(parseToolResult<{ error?: string }>(changed).error).toBe('DRAFT_CHANGED');
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses a token minted for a different target', async () => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const first = parseToolResult<PhaseOne>(
      await h.callTool('outlook_mark_read', { id: 'm1', isRead: true }),
    );
    const other = await h.callTool('outlook_mark_read', {
      id: 'm2',
      isRead: true,
      confirmToken: first.confirmToken,
    });
    expect(other.isError).toBe(true);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('writes when a client that can be prompted accepts', async () => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client), {
      elicitation: async () => ({ action: 'accept', content: { confirmed: true } }),
    });
    const res = await h.callTool('outlook_move_message', { id: 'm1', destination: 'archive' });
    expect(res.isError).toBeFalsy();
    expect(writes()).toHaveLength(1);
    expect(parseToolResult<{ moved?: boolean }>(res).moved).toBe(true);
    await h.close();
  });

  it('writes nothing when a client that can be prompted declines', async () => {
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client), {
      elicitation: async () => ({ action: 'decline' }),
    });
    await h.callTool('outlook_send_mail', { to: ['a@example.com'], subject: 's', body: 'b' });
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses outright under MCP_CONFIRM_MODE=refuse', async () => {
    process.env.MCP_CONFIRM_MODE = 'refuse';
    const { client, writes } = stubClient();
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const res = parseToolResult<{ reason?: string; confirmToken?: string }>(
      await h.callTool('outlook_send_mail', { to: ['a@example.com'], subject: 's', body: 'b' }),
    );
    expect(res.reason).toBe('confirmation-unsupported');
    expect(res.confirmToken).toBeUndefined();
    expect(writes()).toHaveLength(0);
    await h.close();
  });
});
