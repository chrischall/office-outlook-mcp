import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { buildQueryString } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_get_unread` — the one-call triage batch: the unread list, each
 * message's text body, and for an invite the event it refers to. Fixtures are
 * synthetic (example.test addresses, placeholder text).
 */

const wire = (path: string, opts?: { query?: Record<string, unknown> }) =>
  `${path}${buildQueryString((opts?.query ?? {}) as Record<string, string | number | undefined>)}`;

const who = (name: string) => ({ EmailAddress: { Name: name, Address: `${name.toLowerCase()}@example.test` } });

const MAIL = {
  '@odata.type': '#Microsoft.OutlookServices.Message',
  Id: 'm-mail',
  Subject: 'Plain note',
  From: who('Alice'),
  ToRecipients: [who('Me')],
  CcRecipients: [who('Bob')],
  ReceivedDateTime: '2026-10-10T09:00:00Z',
  Importance: 'High',
  HasAttachments: true,
  Categories: ['Blue'],
  Flag: { FlagStatus: 'Flagged' },
  ConversationId: 'conv-1',
  IsRead: false,
};
const REQUEST = {
  '@odata.type': '#Microsoft.OutlookServices.EventMessage',
  Id: 'm-req',
  Subject: 'Invite',
  From: who('Carol'),
  MeetingMessageType: 'MeetingRequest',
};
const CANCELLED = { ...REQUEST, Id: 'm-cxl', MeetingMessageType: 'MeetingCancelled' };
const RESPONSE = { ...REQUEST, Id: 'm-rsp', MeetingMessageType: 'MeetingTenativelyAccepted' };
const EVENT = {
  Id: 'ev-1',
  Subject: 'Sync',
  Start: { DateTime: '2026-10-12T15:00:00.0000000', TimeZone: 'UTC' },
  End: { DateTime: '2026-10-12T15:30:00.0000000', TimeZone: 'UTC' },
  Location: { DisplayName: 'Room 1' },
  Organizer: who('Carol'),
  ResponseStatus: { Response: 'NotResponded', Time: '0001-01-01T00:00:00Z' },
  IsCancelled: false,
};

function stub(list: Record<string, unknown>[], opts: { body?: string; failId?: string; nextLink?: string } = {}) {
  const paths: string[] = [];
  let inFlight = 0;
  let peak = 0;
  const get = vi.fn(async (p: string, o?: { query?: Record<string, unknown>; text?: boolean }) => {
    const path = wire(p, o);
    paths.push(path);
    if (p.includes('/mailfolders/')) {
      return { value: list, ...(opts.nextLink ? { '@odata.nextLink': opts.nextLink } : {}) };
    }
    inFlight++;
    peak = Math.max(peak, inFlight);
    await new Promise((r) => setTimeout(r, 5));
    inFlight--;
    const id = decodeURIComponent(p.split('/me/messages/')[1] ?? '');
    if (id === opts.failId) throw new Error('boom');
    const out: Record<string, unknown> = { Id: id, Body: { ContentType: 'Text', Content: opts.body ?? `  body of ${id}  ` } };
    if (decodeURIComponent(path).includes('EventMessage/Event')) out.Event = EVENT;
    return out;
  });
  const getAbsolute = vi.fn(async () => ({ value: [] }));
  const write = vi.fn(async () => ({}));
  const client = { get, getAbsolute, write } as unknown as OutlookClient;
  return { client, get, getAbsolute, write, paths, peak: () => peak };
}

async function run(client: OutlookClient, args: Record<string, unknown> = {}) {
  const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
  const res = await h.callTool('outlook_get_unread', args);
  await h.close();
  return res;
}

type Out = { count: number; items: Record<string, unknown>[]; nextLink?: string; untrusted_content?: boolean };

describe('outlook_get_unread', () => {
  it('lists unread inbox mail newest first in one call, selecting the meeting type', async () => {
    const s = stub([]);
    await run(s.client);
    expect(s.paths).toHaveLength(1);
    const p = decodeURIComponent(s.paths[0]);
    expect(p).toContain('/me/mailfolders/inbox/messages');
    expect(p).toContain('$filter=IsRead eq false');
    expect(p).toContain('$orderby=ReceivedDateTime desc');
    expect(p).toContain('$top=25');
    expect(p).toContain('MeetingMessageType');
    expect(p).toContain('Flag');
    expect(p).toContain('CcRecipients');
  });

  it('scopes to a folder and a recent window, ReceivedDateTime first in the filter', async () => {
    vi.useFakeTimers({ now: new Date('2026-10-10T12:00:00Z'), toFake: ['Date'] });
    try {
      const s = stub([]);
      await run(s.client, { folder: 'archive', sinceHours: 24, limit: 10 });
      const p = decodeURIComponent(s.paths[0]);
      expect(p).toContain('/me/mailfolders/archive/messages');
      // Outlook wants $orderby properties to lead the $filter.
      expect(p).toContain('$filter=ReceivedDateTime ge 2026-10-09T12:00:00Z and IsRead eq false');
      expect(p).toContain('$top=10');
    } finally {
      vi.useRealTimers();
    }
  });

  it('rejects a limit above 50', async () => {
    const s = stub([]);
    const res = await run(s.client, { limit: 51 });
    expect(res.isError).toBe(true);
    expect(s.paths).toHaveLength(0);
  });

  it('returns each message projected with its text body, inside the untrusted envelope', async () => {
    const s = stub([MAIL]);
    const res = await run(s.client);
    const out = parseToolResult<Out>(res);
    expect(out.untrusted_content).toBe(true);
    expect(out.count).toBe(1);
    expect(out.items[0]).toEqual({
      id: 'm-mail',
      kind: 'mail',
      receivedAt: '2026-10-10T09:00:00Z',
      from: 'Alice <alice@example.test>',
      to: ['Me <me@example.test>'],
      cc: ['Bob <bob@example.test>'],
      subject: 'Plain note',
      importance: 'High',
      hasAttachments: true,
      categories: ['Blue'],
      flag: 'Flagged',
      conversationId: 'conv-1',
      body: 'body of m-mail',
    });
    // The body is fetched as plain text.
    expect(s.get.mock.calls[1][1]).toMatchObject({ text: true });
  });

  it('truncates a long body to maxBodyChars and says so', async () => {
    const s = stub([MAIL], { body: 'x'.repeat(500) });
    const out = parseToolResult<Out>(await run(s.client, { maxBodyChars: 100 }));
    expect(out.items[0].body).toBe('x'.repeat(100));
    expect(out.items[0].bodyTruncated).toBe(true);
  });

  it('defaults the truncation to 4000 characters', async () => {
    const s = stub([MAIL], { body: 'y'.repeat(5000) });
    const out = parseToolResult<Out>(await run(s.client));
    expect(String(out.items[0].body)).toHaveLength(4000);
    expect(out.items[0].bodyTruncated).toBe(true);
  });

  it('skips the per-message fetch for plain mail when includeBody is false', async () => {
    const s = stub([MAIL]);
    const out = parseToolResult<Out>(await run(s.client, { includeBody: false }));
    expect(s.paths).toHaveLength(1);
    expect(out.items[0]).not.toHaveProperty('body');
  });

  it('classifies invites and resolves the event for requests and cancellations', async () => {
    const s = stub([REQUEST, CANCELLED, RESPONSE, MAIL]);
    const out = parseToolResult<Out>(await run(s.client));
    expect(out.items.map((i) => i.kind)).toEqual([
      'meetingRequest',
      'meetingCancelled',
      'meetingResponse',
      'mail',
    ]);
    expect(out.items[0].event).toEqual({
      id: 'ev-1',
      subject: 'Sync',
      start: '2026-10-12T15:00:00.0000000',
      end: '2026-10-12T15:30:00.0000000',
      timeZone: 'UTC',
      location: 'Room 1',
      organizer: 'Carol <carol@example.test>',
      responseStatus: 'NotResponded',
      isCancelled: false,
    });
    expect(out.items[1].event).toBeDefined();
    // A response to my own invite carries no event to act on.
    expect(out.items[2]).not.toHaveProperty('event');
    const expanded = s.paths.filter((p) => decodeURIComponent(p).includes('$expand=Microsoft.OutlookServices.EventMessage/Event'));
    expect(expanded.map((p) => p.split('?')[0])).toEqual(['/me/messages/m-req', '/me/messages/m-cxl']);
  });

  it('still resolves an invite event when bodies are off', async () => {
    const s = stub([REQUEST, MAIL]);
    const out = parseToolResult<Out>(await run(s.client, { includeBody: false }));
    expect(s.paths).toHaveLength(2);
    expect(out.items[0].event).toMatchObject({ id: 'ev-1' });
    expect(out.items[0]).not.toHaveProperty('body');
  });

  it('falls back to @odata.type when MeetingMessageType is absent', async () => {
    const bare = { '@odata.type': '#Microsoft.OutlookServices.EventMessageRequest', Id: 'm-x' };
    const s = stub([bare]);
    const out = parseToolResult<Out>(await run(s.client, { includeBody: false }));
    expect(out.items[0].kind).toBe('meetingRequest');
  });

  it('fetches at most four messages at once', async () => {
    const many = Array.from({ length: 10 }, (_, i) => ({ ...MAIL, Id: `m${i}` }));
    const s = stub(many);
    const out = parseToolResult<Out>(await run(s.client));
    expect(out.count).toBe(10);
    expect(s.peak()).toBeLessThanOrEqual(4);
    expect(s.peak()).toBeGreaterThan(1);
  });

  it('keeps the batch when one message fails to load, flagging that item', async () => {
    const s = stub([MAIL, { ...MAIL, Id: 'bad' }], { failId: 'bad' });
    const out = parseToolResult<Out>(await run(s.client));
    expect(out.count).toBe(2);
    expect(out.items[0].body).toBe('body of m-mail');
    expect(out.items[1]).not.toHaveProperty('body');
    expect(String(out.items[1].error)).toMatch(/boom/);
  });

  it('passes the paging link through and follows one back', async () => {
    const link = 'https://outlook.office.com/api/v2.0/me/mailfolders/inbox/messages?$skip=25';
    const s = stub([MAIL], { nextLink: link });
    const out = parseToolResult<Out>(await run(s.client, { includeBody: false }));
    expect(out.nextLink).toBe(link);
    await run(s.client, { nextLink: link });
    expect(s.getAbsolute).toHaveBeenCalledWith(link, expect.anything());
  });

  it('never marks anything read and says how to', async () => {
    const s = stub([MAIL, REQUEST]);
    await run(s.client);
    expect(s.write).not.toHaveBeenCalled();
    const h = await createTestHarness((srv: McpServer) => registerMailTools(srv, s.client));
    const tool = (await h.listTools()).find((t) => t.name === 'outlook_get_unread');
    await h.close();
    expect(tool?.description).toMatch(/outlook_mark_read/);
    expect(tool?.description).toMatch(/untrusted/i);
  });
});
