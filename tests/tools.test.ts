import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerCalendarTools } from '../src/tools/calendar.js';
import { registerDirectoryTools } from '../src/tools/directory.js';
import type { OutlookClient } from '../src/client.js';

/** A client stub that records calls and returns canned payloads. */
function stubClient(overrides: Partial<Record<string, unknown>> = {}) {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  const client = {
    get: vi.fn(async (path: string) => {
      calls.push({ method: 'GET', path });
      if (path.includes('$select=IsRead')) return { IsRead: true };
      if (path.startsWith('/me/messages/') && !path.includes('/attachments')) {
        return { Id: 'm1', Subject: 'Hi', Body: { ContentType: 'Text', Content: 'body' } };
      }
      return { value: [] };
    }),
    write: vi.fn(async (method: string, path: string, body?: unknown) => {
      calls.push({ method, path, body });
      return { Id: 'new-1' };
    }),
    ...overrides,
  } as unknown as OutlookClient;
  return { client, calls };
}

const harnessFor = (register: (s: McpServer, c: OutlookClient) => void, client: OutlookClient) =>
  createTestHarness((server: McpServer) => register(server, client));

describe('read tools', () => {
  it('defaults to the inbox, newest first, without fetching bodies', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerMailTools, client);
    await h.callTool('outlook_list_messages', {});
    expect(calls[0].path).toContain('/me/mailfolders/inbox/messages');
    expect(calls[0].path).toContain('ReceivedDateTime%20desc');
    // A listing must not pull Body — that is the whole point of $select here.
    expect(calls[0].path).not.toContain('Body,');
    await h.close();
  });

  it('filters to unread on request', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerMailTools, client);
    await h.callTool('outlook_list_messages', { unreadOnly: true });
    expect(decodeURIComponent(calls[0].path)).toContain('IsRead eq false');
    await h.close();
  });

  it('rejects search combined with unreadOnly rather than sending a doomed query', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerMailTools, client);
    const res = await h.callTool('outlook_list_messages', { search: 'x', unreadOnly: true });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/cannot be combined/);
    // No network call at all — the conflict is caught before the request.
    expect(calls).toHaveLength(0);
    await h.close();
  });

  it('searches the whole mailbox and omits $orderby, which $search rejects', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerMailTools, client);
    await h.callTool('outlook_list_messages', { search: 'quarterly report' });
    expect(calls[0].path).toContain('/me/messages');
    expect(calls[0].path).not.toContain('/mailfolders/');
    expect(calls[0].path).not.toContain('$orderby');
    await h.close();
  });

  it('asks for a text body when reading one message, and HTML only on request', async () => {
    const { client } = stubClient();
    const h = await harnessFor(registerMailTools, client);
    await h.callTool('outlook_get_message', { id: 'm1' });
    expect((client.get as ReturnType<typeof vi.fn>).mock.calls[0][1]).toMatchObject({ text: true });
    await h.callTool('outlook_get_message', { id: 'm1', html: true });
    expect((client.get as ReturnType<typeof vi.fn>).mock.calls[1][1]).toMatchObject({ text: false });
    await h.close();
  });

  it('passes a Windows time zone through to the calendar view', async () => {
    const { client } = stubClient();
    const h = await harnessFor(registerCalendarTools, client);
    await h.callTool('outlook_list_events', {
      start: '2026-09-20T00:00:00Z',
      end: '2026-09-27T00:00:00Z',
      timeZone: 'Eastern Standard Time',
    });
    const [path, opts] = (client.get as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(path).toContain('/me/calendarview');
    expect(opts).toMatchObject({ prefer: 'outlook.timezone="Eastern Standard Time"' });
    await h.close();
  });

  it('registers the expected read surface', async () => {
    const { client } = stubClient();
    const h = await harnessFor(registerDirectoryTools, client);
    const names = (await h.listTools()).map((t) => t.name).sort();
    expect(names).toEqual([
      'outlook_get_mailbox_settings',
      'outlook_get_profile',
      'outlook_list_contacts',
      'outlook_list_people',
      'outlook_list_tasks',
    ]);
    await h.close();
  });
});

describe('write tools are confirm-gated', () => {
  const writeTools = [
    ['outlook_send_mail', { to: ['a@example.com'], subject: 's', body: 'b' }],
    ['outlook_create_draft', { to: ['a@example.com'], subject: 's', body: 'b' }],
    ['outlook_mark_read', { id: 'm1', isRead: true }],
    ['outlook_move_message', { id: 'm1', destination: 'archive' }],
    ['outlook_create_event', { subject: 's', start: '2026-09-22T15:00:00', end: '2026-09-22T16:00:00' }],
  ] as const;

  it.each(writeTools)('%s makes no network call without confirm', async (name, args) => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerWriteTools, client);
    const res = parseToolResult<{ dryRun?: boolean }>(await h.callTool(name, { ...args }));
    expect(res.dryRun).toBe(true);
    expect(calls).toHaveLength(0);
    await h.close();
  });

  it('sends only when confirmed, and reports the recipients it used', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerWriteTools, client);
    const res = parseToolResult<{ sent?: boolean; recipients?: string }>(
      await h.callTool('outlook_send_mail', {
        to: ['a@example.com'],
        cc: ['b@example.com'],
        subject: 'Subj',
        body: 'Body',
        confirm: true,
      }),
    );
    expect(res.sent).toBe(true);
    expect(res.recipients).toBe('a@example.com, b@example.com');
    const sent = calls.find((c) => c.path === '/me/sendmail');
    expect(sent?.method).toBe('POST');
    expect(sent?.body).toMatchObject({
      Message: {
        Subject: 'Subj',
        Body: { ContentType: 'Text' },
        ToRecipients: [{ EmailAddress: { Address: 'a@example.com' } }],
        CcRecipients: [{ EmailAddress: { Address: 'b@example.com' } }],
      },
      SaveToSentItems: true,
    });
    await h.close();
  });

  it('verifies a mark-read by re-reading rather than trusting the status', async () => {
    const { client, calls } = stubClient();
    const h = await harnessFor(registerWriteTools, client);
    const res = parseToolResult<{ updated?: boolean }>(
      await h.callTool('outlook_mark_read', { id: 'm1', isRead: true, confirm: true }),
    );
    expect(res.updated).toBe(true);
    // PATCH, then a GET that re-reads the field the write requested.
    expect(calls.map((c) => c.method)).toEqual(['PATCH', 'GET']);
    expect(calls[1].path).toContain('$select=IsRead');
    await h.close();
  });

  it('warns when the write was accepted but the value did not move', async () => {
    // The false-green case: a 2xx with no observable change must not report success.
    const { client } = stubClient({
      get: vi.fn(async () => ({ IsRead: false })),
    });
    const h = await harnessFor(registerWriteTools, client);
    const res = parseToolResult<{ updated?: boolean; warning?: string }>(
      await h.callTool('outlook_mark_read', { id: 'm1', isRead: true, confirm: true }),
    );
    expect(res.updated).toBe(false);
    expect(res.warning).toMatch(/did not change/);
    await h.close();
  });
});
