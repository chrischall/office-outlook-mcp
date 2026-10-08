import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import { fullEvent } from '../src/view.js';
import type { OutlookClient } from '../src/client.js';

/**
 * Field names and shapes are from live events read 2026-10-08: a Teams
 * meeting carries `IsOnlineMeeting: true`, `OnlineMeetingProvider:
 * "TeamsForBusiness"` and `OnlineMeeting.JoinUrl`, while the legacy
 * `OnlineMeetingUrl` is an empty string. A non-online meeting has
 * `OnlineMeeting: {}`.
 */
const JOIN = 'https://teams.microsoft.com/l/meetup-join/abc';

type Ev = Record<string, unknown>;

function baseEvent(over: Ev = {}): Ev {
  return {
    Id: 'e1',
    Subject: 'Sync',
    IsOrganizer: true,
    Start: { DateTime: '2026-10-12T10:00:00.0000000', TimeZone: 'Eastern Standard Time' },
    End: { DateTime: '2026-10-12T10:30:00.0000000', TimeZone: 'Eastern Standard Time' },
    Attendees: [
      { Type: 'Required', EmailAddress: { Address: 'a@x.test', Name: 'A' }, Status: { Response: 'Accepted' } },
      { Type: 'Optional', EmailAddress: { Address: 'b@x.test' } },
    ],
    IsOnlineMeeting: false,
    OnlineMeetingProvider: 'Unknown',
    OnlineMeeting: {},
    ...over,
  };
}

/**
 * A stateful stub: PATCHes merge into the stored event so the re-read sees
 * what Outlook would, and `persist: false` simulates a write that 2xx's but
 * does not stick.
 */
function stub(event: Ev, opts: { persist?: boolean; created?: Ev } = {}) {
  const calls: { method: string; path: string; body?: unknown; prefer?: string }[] = [];
  let stored = { ...event };
  const client = {
    get: vi.fn(async (path: string, o?: { prefer?: string }) => {
      calls.push({ method: 'GET', path, prefer: o?.prefer });
      if (path.includes('MailboxSettings')) return { TimeZone: 'Eastern Standard Time' };
      if (path.startsWith('/me/events/')) return stored;
      return {};
    }),
    write: vi.fn(async (method: string, path: string, body?: unknown) => {
      calls.push({ method, path, body });
      if (method === 'PATCH' && opts.persist !== false) {
        const b = body as Ev;
        stored = { ...stored, ...b };
        if (b.IsOnlineMeeting === true) stored.OnlineMeeting = { JoinUrl: JOIN };
      }
      if (method === 'POST' && path === '/me/events') {
        stored = { Id: 'new-1', ...(opts.created ?? { OnlineMeeting: { JoinUrl: JOIN } }) };
        return stored;
      }
      return stored;
    }),
  } as unknown as OutlookClient;
  return { client, calls };
}

async function harness(client: OutlookClient) {
  return createTestHarness((s: McpServer) => registerWriteTools(s, client));
}

async function confirmed(h: TestHarness, name: string, args: Record<string, unknown>) {
  const first = parseToolResult<{ confirmToken?: string }>(await h.callTool(name, args));
  return h.callTool(name, { ...args, confirmToken: first.confirmToken });
}

const writes = <T extends { method: string }>(calls: T[]): T[] => calls.filter((c) => c.method !== 'GET');

describe('outlook_create_event Teams meetings', () => {
  const args = {
    subject: 'Planning',
    start: '2026-10-12T15:00:00',
    end: '2026-10-12T15:30:00',
    attendees: ['a@x.test'],
    optionalAttendees: ['b@x.test'],
  };

  it('attaches a Teams meeting by default and marks optional attendees', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    await confirmed(h, 'outlook_create_event', args);
    const post = writes(calls)[0];
    expect(post.path).toBe('/me/events');
    expect(post.body).toMatchObject({
      IsOnlineMeeting: true,
      OnlineMeetingProvider: 'TeamsForBusiness',
      Attendees: [
        { Type: 'Required', EmailAddress: { Address: 'a@x.test' } },
        { Type: 'Optional', EmailAddress: { Address: 'b@x.test' } },
      ],
    });
    await h.close();
  });

  it('leaves Teams off only when asked', async () => {
    const { client, calls } = stub(baseEvent(), { created: { OnlineMeeting: {} } });
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, 'outlook_create_event', { ...args, teamsMeeting: false }));
    expect(writes(calls)[0].body).not.toHaveProperty('IsOnlineMeeting');
    expect(res).not.toHaveProperty('warning');
    await h.close();
  });

  it('returns the Teams join link', async () => {
    const { client } = stub(baseEvent());
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, 'outlook_create_event', args));
    expect(res).toMatchObject({ created: true, Id: 'new-1', JoinUrl: JOIN });
    await h.close();
  });

  it('warns when Outlook created the event without the Teams link it was asked for', async () => {
    // A 201 is not proof the Teams meeting was provisioned; the link is.
    const { client } = stub(baseEvent(), { created: { OnlineMeeting: {} } });
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, 'outlook_create_event', args));
    expect(res.created).toBe(true);
    expect(String(res.warning)).toMatch(/Teams/);
    await h.close();
  });
});

describe('outlook_update_event', () => {
  it('writes nothing without a confirmToken', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    const res = parseToolResult<{ status?: string }>(
      await h.callTool('outlook_update_event', { id: 'e1', subject: 'New' }),
    );
    expect(res.status).toBe('confirmation-required');
    expect(writes(calls)).toHaveLength(0);
    await h.close();
  });

  it('refuses to edit a meeting someone else organizes', async () => {
    // An attendee's PATCH changes only their own copy; the organizer and the
    // other attendees never see it. Saying "updated" would be a lie.
    const { client, calls } = stub(baseEvent({ IsOrganizer: false }));
    const h = await harness(client);
    const res = await h.callTool('outlook_update_event', { id: 'e1', subject: 'New' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res.content)).toMatch(/organizer/i);
    expect(writes(calls)).toHaveLength(0);
    await h.close();
  });

  it('reads the current event in the mailbox zone', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    await h.callTool('outlook_update_event', { id: 'e1', subject: 'New' });
    const read = calls.find((c) => c.path.startsWith('/me/events/e1'));
    expect(read?.prefer).toBe('outlook.timezone="Eastern Standard Time"');
    await h.close();
  });

  it('keeps the duration when only the start moves', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    await confirmed(h, 'outlook_update_event', { id: 'e1', start: '2026-10-13T15:00:00' });
    const patch = writes(calls)[0];
    expect(patch.method).toBe('PATCH');
    expect(patch.path).toBe('/me/events/e1');
    expect(patch.body).toMatchObject({
      Start: { DateTime: '2026-10-13T15:00:00', TimeZone: 'Eastern Standard Time' },
      End: { DateTime: '2026-10-13T15:30:00', TimeZone: 'Eastern Standard Time' },
    });
    await h.close();
  });

  it('adds and removes attendees against the current list, keeping their types', async () => {
    const { client, calls } = stub(baseEvent({ IsOnlineMeeting: true, OnlineMeetingProvider: 'TeamsForBusiness', OnlineMeeting: { JoinUrl: JOIN } }));
    const h = await harness(client);
    await confirmed(h, 'outlook_update_event', {
      id: 'e1',
      addAttendees: ['c@x.test', 'A@X.test'],
      addOptionalAttendees: ['d@x.test'],
      removeAttendees: ['B@x.test'],
    });
    const body = writes(calls)[0].body as Ev;
    expect(body.Attendees).toEqual([
      { Type: 'Required', EmailAddress: { Address: 'a@x.test', Name: 'A' } },
      { Type: 'Required', EmailAddress: { Address: 'c@x.test' } },
      { Type: 'Optional', EmailAddress: { Address: 'd@x.test' } },
    ]);
    // Already a Teams meeting: nothing to add.
    expect(body).not.toHaveProperty('IsOnlineMeeting');
    await h.close();
  });

  it('adds a Teams meeting to an event that lacks one', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, 'outlook_update_event', { id: 'e1', subject: 'Renamed' }));
    expect(writes(calls)[0].body).toMatchObject({
      Subject: 'Renamed',
      IsOnlineMeeting: true,
      OnlineMeetingProvider: 'TeamsForBusiness',
    });
    expect(res).toMatchObject({ updated: true, JoinUrl: JOIN });
    await h.close();
  });

  it('adds only the Teams meeting when that is the one change', async () => {
    const { client, calls } = stub(baseEvent());
    const h = await harness(client);
    await confirmed(h, 'outlook_update_event', { id: 'e1' });
    expect(writes(calls)[0].body).toEqual({ IsOnlineMeeting: true, OnlineMeetingProvider: 'TeamsForBusiness' });
    await h.close();
  });

  it('refuses an update that changes nothing', async () => {
    const { client, calls } = stub(baseEvent({ IsOnlineMeeting: true, OnlineMeetingProvider: 'TeamsForBusiness', OnlineMeeting: { JoinUrl: JOIN } }));
    const h = await harness(client);
    const res = await h.callTool('outlook_update_event', { id: 'e1' });
    expect(res.isError).toBe(true);
    expect(writes(calls)).toHaveLength(0);
    await h.close();
  });

  it('verifies by re-reading and warns when a change did not stick', async () => {
    const { client } = stub(baseEvent(), { persist: false });
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, 'outlook_update_event', { id: 'e1', subject: 'Renamed' }));
    expect(res.updated).toBe(false);
    expect(String(res.warning)).toMatch(/Subject/);
    expect(String(res.warning)).toMatch(/Teams/);
    await h.close();
  });
});

describe('event views', () => {
  it('shows the Teams join link, not the legacy empty OnlineMeetingUrl', () => {
    const v = fullEvent({ ...baseEvent(), OnlineMeetingUrl: '', OnlineMeeting: { JoinUrl: JOIN } } as never);
    expect(v.JoinUrl).toBe(JOIN);
  });
});
