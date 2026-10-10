import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_delete_event` — remove one event, one occurrence, or (by its
 * SeriesMaster id) a whole series. Only the organizer may delete a meeting;
 * when it has attendees they are sent a cancellation, so that case is gated.
 * Fixtures are synthetic (example.test addresses, placeholder subjects).
 */

type Ev = Record<string, unknown>;

const attendees = (n: number) =>
  Array.from({ length: n }, (_, i) => ({
    Type: 'Required',
    EmailAddress: { Address: `person${i}@example.test` },
    Status: { Response: 'Accepted' },
  }));

function meeting(over: Ev = {}): Ev {
  return {
    Id: 'ev-1',
    Subject: 'Planning review',
    IsOrganizer: true,
    IsCancelled: false,
    Type: 'SingleInstance',
    Start: { DateTime: '2026-10-12T15:00:00.0000000', TimeZone: 'Eastern Standard Time' },
    End: { DateTime: '2026-10-12T15:30:00.0000000', TimeZone: 'Eastern Standard Time' },
    Organizer: { EmailAddress: { Name: 'Me', Address: 'me@example.test' } },
    ResponseStatus: { Response: 'Organizer' },
    Attendees: attendees(3),
    ...over,
  };
}

/**
 * A stateful stub: a DELETE removes the stored event so the read-back 404s,
 * as Outlook's does. `persist: false` simulates a 204 that did not stick.
 */
function stub(
  event: Ev | undefined,
  opts: { persist?: boolean; writeError?: Error; readBackError?: Error } = {},
) {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  let stored = event ? { ...event } : undefined;
  let deleted = false;
  const get = vi.fn(async (path: string) => {
    calls.push({ method: 'GET', path });
    if (path.includes('MailboxSettings')) return { TimeZone: 'Eastern Standard Time' };
    if (path.startsWith('/me/events/')) {
      if (deleted && opts.readBackError) throw opts.readBackError;
      if (!stored) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      return stored;
    }
    return {};
  });
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    if (opts.writeError) throw opts.writeError;
    deleted = true;
    if (opts.persist !== false) stored = undefined;
    return undefined;
  });
  const client = { get, write } as unknown as OutlookClient;
  const writes = () => calls.filter((c) => c.method !== 'GET');
  return { client, calls, writes };
}

async function harness(client: OutlookClient) {
  return createTestHarness((s: McpServer) => registerWriteTools(s, client));
}

async function confirmed(h: TestHarness, args: Record<string, unknown>) {
  const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_delete_event', args));
  return h.callTool('outlook_delete_event', { ...args, confirmToken: first.confirmToken });
}

describe('outlook_delete_event', () => {
  it('previews subject, time, attendee count and scope, and deletes nothing without a token', async () => {
    const { client, writes } = stub(meeting());
    const h = await harness(client);
    const res = parseToolResult<{ status?: string; preview?: Record<string, unknown> }>(
      await h.callTool('outlook_delete_event', { id: 'ev-1' }),
    );
    expect(res.status).toBe('confirmation-required');
    expect(res.preview).toMatchObject({ method: 'DELETE', path: '/me/events/ev-1' });
    const text = JSON.stringify(res);
    expect(text).toMatch(/Planning review/);
    expect(text).toMatch(/2026-10-12T15:00/);
    expect(text).toMatch(/3 attendees/);
    expect(text).toMatch(/cancellation/i);
    expect(text).toMatch(/this event/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('deletes with the token and verifies by re-reading (404 = gone)', async () => {
    const { client, writes, calls } = stub(meeting());
    const h = await harness(client);
    const res = await confirmed(h, { id: 'ev-1' });
    expect(res.isError).toBeFalsy();
    expect(writes()).toEqual([{ method: 'DELETE', path: '/me/events/ev-1', body: undefined }]);
    expect(parseToolResult<Ev>(res)).toMatchObject({
      deleted: true,
      Id: 'ev-1',
      cancellationSent: true,
    });
    const last = calls[calls.length - 1];
    expect(last.method).toBe('GET');
    expect(last.path).toMatch(/^\/me\/events\/ev-1/);
    await h.close();
  });

  it('says a SeriesMaster id deletes the whole series', async () => {
    const { client } = stub(meeting({ Type: 'SeriesMaster' }));
    const h = await harness(client);
    const res = parseToolResult<{ preview?: Record<string, unknown> }>(
      await h.callTool('outlook_delete_event', { id: 'ev-1' }),
    );
    expect(JSON.stringify(res.preview)).toMatch(/whole series/i);
    await h.close();
  });

  it.each(['Occurrence', 'Exception'])('says an %s id deletes only that occurrence', async (Type) => {
    const { client } = stub(meeting({ Type, SeriesMasterId: 'master-1' }));
    const h = await harness(client);
    const res = parseToolResult<{ preview?: Record<string, unknown> }>(
      await h.callTool('outlook_delete_event', { id: 'ev-1' }),
    );
    const text = JSON.stringify(res.preview);
    expect(text).toMatch(/only this occurrence/i);
    expect(text).not.toMatch(/whole series/i);
    await h.close();
  });

  it("refuses a meeting you do not organize, pointing at a decline instead", async () => {
    const { client, writes } = stub(
      meeting({ IsOrganizer: false, ResponseStatus: { Response: 'Accepted' } }),
    );
    const h = await harness(client);
    const res = await h.callTool('outlook_delete_event', { id: 'ev-1' });
    expect(res.isError).toBe(true);
    const text = JSON.stringify(res);
    expect(text).toMatch(/organiz/i);
    expect(text).toMatch(/outlook_respond_to_invite/);
    expect(text).toMatch(/decline/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('removes a meeting the organizer already cancelled without asking: nobody is told', async () => {
    // An attendee's copy of a cancelled meeting stays on the calendar until
    // removed, and outlook_respond_to_invite refuses cancelled meetings — this
    // is the only way to clear it, and it reaches no one.
    const { client, writes } = stub(
      meeting({ IsOrganizer: false, IsCancelled: true, ResponseStatus: { Response: 'Accepted' } }),
    );
    const h = await harness(client);
    const res = await h.callTool('outlook_delete_event', { id: 'ev-1' });
    expect(res.isError).toBeFalsy();
    expect(writes()).toEqual([{ method: 'DELETE', path: '/me/events/ev-1', body: undefined }]);
    expect(parseToolResult<Ev>(res)).toMatchObject({ deleted: true, cancellationSent: false });
    await h.close();
  });

  it('deletes your own appointment with no attendees without asking: only your calendar changes', async () => {
    const { client, writes } = stub(meeting({ Attendees: [] }));
    const h = await harness(client);
    const res = await h.callTool('outlook_delete_event', { id: 'ev-1' });
    expect(res.isError).toBeFalsy();
    expect(writes()).toEqual([{ method: 'DELETE', path: '/me/events/ev-1', body: undefined }]);
    expect(parseToolResult<Ev>(res)).toMatchObject({ deleted: true, cancellationSent: false });
    await h.close();
  });

  it('warns when Outlook accepted the delete but the event is still there', async () => {
    const { client } = stub(meeting(), { persist: false });
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, { id: 'ev-1' }));
    expect(res.deleted).toBe(false);
    expect(String(res.warning)).toMatch(/still/i);
    await h.close();
  });

  it('reports a failed read-back that is not a 404 as unverified', async () => {
    const { client } = stub(meeting(), { readBackError: new ApiError(503, 'Outlook 503: busy') });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'ev-1' });
    expect(res.isError).toBeFalsy();
    const body = parseToolResult<Ev>(res);
    expect(body.deleted).toBeNull();
    expect(String(body.warning)).toMatch(/unverified/);
    expect(String(body.warning)).toMatch(/503/);
    await h.close();
  });

  it('reports an unconfirmed delete as unknown instead of an error', async () => {
    const { client } = stub(meeting(), {
      writeError: new WriteOutcomeUnknownError('Outlook', 'DELETE', { timeoutMs: 60_000 }),
    });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'ev-1' });
    expect(res.isError).toBeFalsy();
    expect(parseToolResult<Ev>(res).status).toBe('unknown');
    await h.close();
  });

  it('surfaces an event that does not exist without writing', async () => {
    const { client, writes } = stub(undefined);
    const h = await harness(client);
    const res = await h.callTool('outlook_delete_event', { id: 'nope' });
    expect(res.isError).toBe(true);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses a token minted for a different event', async () => {
    const { client, writes } = stub(meeting());
    const h = await harness(client);
    const first = parseToolResult<{ confirmToken?: string }>(
      await h.callTool('outlook_delete_event', { id: 'ev-1' }),
    );
    const other = await h.callTool('outlook_delete_event', { id: 'ev-2', confirmToken: first.confirmToken });
    expect(other.isError).toBe(true);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('is annotated destructive', async () => {
    const { client } = stub(meeting());
    const h = await harness(client);
    const tool = (await h.client.listTools()).tools.find((t) => t.name === 'outlook_delete_event');
    expect(tool?.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    });
    await h.close();
  });
});
