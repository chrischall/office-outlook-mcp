import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, buildQueryString, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_respond_to_invite` — accept / tentatively accept / decline a
 * meeting, by event id or by the invite message it arrived as. Fixtures are
 * synthetic (example.test addresses, placeholder subjects).
 */

type Ev = Record<string, unknown>;

const wire = (path: string, opts?: { query?: Record<string, unknown> }) =>
  `${path}${buildQueryString((opts?.query ?? {}) as Record<string, string | number | undefined>)}`;

function invite(over: Ev = {}): Ev {
  return {
    Id: 'ev-1',
    Subject: 'Quarterly sync',
    IsOrganizer: false,
    IsCancelled: false,
    Start: { DateTime: '2026-10-12T15:00:00.0000000', TimeZone: 'UTC' },
    End: { DateTime: '2026-10-12T15:30:00.0000000', TimeZone: 'UTC' },
    Organizer: { EmailAddress: { Name: 'Carol', Address: 'carol@example.test' } },
    ResponseStatus: { Response: 'NotResponded' },
    ...over,
  };
}

const RESPONSE_OF: Record<string, string> = {
  accept: 'Accepted',
  tentativelyaccept: 'TentativelyAccepted',
  decline: 'Declined',
};

/**
 * A stateful stub: a response POST updates the stored ResponseStatus so the
 * read-back sees what Outlook would; `persist: false` simulates a 202 that
 * does not stick, `goneAfterDecline` a declined event Outlook removed.
 */
function stub(
  event: Ev | undefined,
  opts: {
    persist?: boolean;
    goneAfterDecline?: boolean;
    writeError?: Error;
    noEventOnMessage?: boolean;
    /** The invite message's record, replacing the default `{Id, Event}`. */
    message?: Ev;
    /** The read-back after the response POST throws this. */
    readBackError?: Error;
  } = {},
) {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  let stored = event ? { ...event } : undefined;
  let declined = false;
  let responded = false;
  const get = vi.fn(async (p: string, o?: { query?: Record<string, unknown> }) => {
    const path = wire(p, o);
    calls.push({ method: 'GET', path });
    if (p.startsWith('/me/messages/')) {
      if (opts.message) return opts.message;
      return opts.noEventOnMessage ? { Id: 'm-1' } : { Id: 'm-1', Event: stored };
    }
    if (p.startsWith('/me/events/')) {
      if (declined && opts.goneAfterDecline) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      if (responded && opts.readBackError) throw opts.readBackError;
      return stored;
    }
    return {};
  });
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    if (opts.writeError) throw opts.writeError;
    const verb = path.split('/').pop() ?? '';
    responded = true;
    if (verb === 'decline') declined = true;
    if (opts.persist !== false && stored) stored = { ...stored, ResponseStatus: { Response: RESPONSE_OF[verb] } };
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
  const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_respond_to_invite', args));
  return h.callTool('outlook_respond_to_invite', { ...args, confirmToken: first.confirmToken });
}

describe('outlook_respond_to_invite', () => {
  it('previews subject, time, organizer, response and comment, and writes nothing without a token', async () => {
    const { client, writes } = stub(invite());
    const h = await harness(client);
    const res = parseToolResult<{ status?: string; preview?: Record<string, unknown> }>(
      await h.callTool('outlook_respond_to_invite', { eventId: 'ev-1', response: 'accept', comment: 'See you there' }),
    );
    expect(res.status).toBe('confirmation-required');
    const text = JSON.stringify(res);
    expect(text).toMatch(/Quarterly sync/);
    expect(text).toMatch(/2026-10-12T15:00/);
    expect(text).toMatch(/carol@example\.test/);
    expect(text).toMatch(/accept/i);
    expect(text).toMatch(/See you there/);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it.each([
    ['accept', 'accept', 'Accepted'],
    ['tentative', 'tentativelyaccept', 'TentativelyAccepted'],
    ['decline', 'decline', 'Declined'],
  ])('%s posts to /%s and verifies the response by re-reading', async (response, verb, status) => {
    const { client, writes, calls } = stub(invite());
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, { eventId: 'ev-1', response, comment: 'ok' }));
    expect(writes()).toEqual([
      { method: 'POST', path: `/me/events/ev-1/${verb}`, body: { Comment: 'ok', SendResponse: true } },
    ]);
    expect(res).toMatchObject({ responded: true, eventId: 'ev-1', response: status, sentResponse: true });
    expect(res).not.toHaveProperty('warning');
    // The read-back happens after the write.
    const last = calls[calls.length - 1];
    expect(last.method).toBe('GET');
    expect(last.path).toMatch(/^\/me\/events\/ev-1/);
    await h.close();
  });

  it('resolves a messageId to its event through the EventMessage expand', async () => {
    const { client, writes, calls } = stub(invite());
    const h = await harness(client);
    await confirmed(h, { messageId: 'm-1', response: 'accept' });
    const lookup = decodeURIComponent(calls[0].path);
    expect(lookup).toMatch(/^\/me\/messages\/m-1\?/);
    expect(lookup).toContain('$expand=Microsoft.OutlookServices.EventMessage/Event');
    expect(writes()[0]).toEqual({ method: 'POST', path: '/me/events/ev-1/accept', body: { SendResponse: true } });
    await h.close();
  });

  it('refuses a message that is not a meeting invite', async () => {
    const { client, writes } = stub(invite(), { noEventOnMessage: true });
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { messageId: 'm-1', response: 'accept' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/not a meeting invite/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  // Live 2026-10-10: an invite already declined (its event gone) still reads
  // as a MeetingRequest, but its Event expands to an empty object.
  it.each([
    ['an empty Event', { Event: {} }],
    ['no Event at all', {}],
  ])('says a meeting request with %s is no longer on the calendar', async (_label, extra) => {
    const message = {
      '@odata.type': '#Microsoft.OutlookServices.EventMessage',
      Id: 'm-1',
      MeetingMessageType: 'MeetingRequest',
      ...extra,
    };
    const { client, writes, calls } = stub(invite(), { message });
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { messageId: 'm-1', response: 'accept' });
    expect(res.isError).toBe(true);
    const text = JSON.stringify(res);
    expect(text).toMatch(/no longer on your calendar/i);
    expect(text).toMatch(/nothing to respond to/i);
    expect(text).not.toMatch(/not a meeting invite/i);
    expect(decodeURIComponent(calls[0].path)).toContain('MeetingMessageType');
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('still calls a plain message with an empty Event not a meeting invite', async () => {
    const message = { '@odata.type': '#Microsoft.OutlookServices.Message', Id: 'm-1', Event: {} };
    const { client, writes } = stub(invite(), { message });
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { messageId: 'm-1', response: 'accept' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/not a meeting invite/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it.each([
    [{}, /exactly one/i],
    [{ messageId: 'm-1', eventId: 'ev-1' }, /exactly one/i],
  ])('requires exactly one of messageId / eventId (%o)', async (ids, msg) => {
    const { client, writes } = stub(invite());
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { ...ids, response: 'accept' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(msg);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses a cancelled meeting', async () => {
    const { client, writes } = stub(invite({ IsCancelled: true }));
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { eventId: 'ev-1', response: 'accept' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/cancelled/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses a meeting you organize', async () => {
    const { client, writes } = stub(invite({ IsOrganizer: true, ResponseStatus: { Response: 'Organizer' } }));
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', { eventId: 'ev-1', response: 'decline' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/organiz/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('skips the confirm step when no response is sent to the organizer', async () => {
    // sendResponse: false changes only this mailbox's calendar; nobody else is told.
    const { client, writes } = stub(invite());
    const h = await harness(client);
    const res = parseToolResult<Ev>(
      await h.callTool('outlook_respond_to_invite', { eventId: 'ev-1', response: 'tentative', sendResponse: false }),
    );
    expect(writes()).toEqual([
      // No Comment key at all: Outlook rejects SendResponse:false with any
      // Comment, even an empty one ("'SendResponse' must be true when
      // 'Comment' is not null" — live, 2026-10-10).
      { method: 'POST', path: '/me/events/ev-1/tentativelyaccept', body: { SendResponse: false } },
    ]);
    expect(res).toMatchObject({ responded: true, response: 'TentativelyAccepted', sentResponse: false });
    await h.close();
  });

  it('refuses a comment when no response is sent, before writing anything', async () => {
    const { client, writes } = stub(invite());
    const h = await harness(client);
    const res = await h.callTool('outlook_respond_to_invite', {
      eventId: 'ev-1',
      response: 'accept',
      sendResponse: false,
      comment: 'see you there',
    });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/comment/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('omits Comment from a notified response that has none', async () => {
    const { client, writes } = stub(invite());
    const h = await harness(client);
    await confirmed(h, { eventId: 'ev-1', response: 'accept' });
    expect(writes()).toEqual([
      { method: 'POST', path: '/me/events/ev-1/accept', body: { SendResponse: true } },
    ]);
    await h.close();
  });

  it('warns when Outlook accepted the write but the response did not change', async () => {
    const { client } = stub(invite(), { persist: false });
    const h = await harness(client);
    const res = parseToolResult<Ev>(await confirmed(h, { eventId: 'ev-1', response: 'accept' }));
    expect(res.responded).toBe(false);
    expect(res.response).toBe('NotResponded');
    expect(String(res.warning)).toMatch(/did not change/);
    await h.close();
  });

  it('treats a declined event that is gone from the calendar as declined', async () => {
    const { client } = stub(invite(), { goneAfterDecline: true });
    const h = await harness(client);
    const res = await confirmed(h, { eventId: 'ev-1', response: 'decline' });
    expect(res.isError).toBeFalsy();
    const body = parseToolResult<Ev>(res);
    expect(body).toMatchObject({ responded: true, response: 'Declined' });
    expect(String(body.note)).toMatch(/no longer on (your|the) calendar/i);
    await h.close();
  });

  it('does not read a failed read-back after a decline as "gone" unless it was a 404', async () => {
    const { client } = stub(invite(), { readBackError: new ApiError(503, 'Outlook 503: busy') });
    const h = await harness(client);
    const res = await confirmed(h, { eventId: 'ev-1', response: 'decline' });
    expect(res.isError).toBeFalsy();
    const body = parseToolResult<Ev>(res);
    expect(body.responded).toBeNull();
    expect(body).not.toHaveProperty('note');
    expect(String(body.warning)).toMatch(/unverified/);
    expect(String(body.warning)).toMatch(/503/);
    await h.close();
  });

  it.each(['accept', 'tentative'])('reports a failed read-back after %s as unverified', async (response) => {
    const { client } = stub(invite(), { readBackError: new ApiError(404, 'Outlook 404: ErrorItemNotFound') });
    const h = await harness(client);
    const res = await confirmed(h, { eventId: 'ev-1', response });
    expect(res.isError).toBeFalsy();
    const body = parseToolResult<Ev>(res);
    expect(body).toMatchObject({ responded: null, eventId: 'ev-1', sentResponse: true });
    expect(String(body.warning)).toMatch(/unverified/);
    await h.close();
  });

  it('reports an unconfirmed write as unknown instead of an error', async () => {
    const { client } = stub(invite(), { writeError: new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await confirmed(h, { eventId: 'ev-1', response: 'accept' });
    expect(res.isError).toBeFalsy();
    expect(parseToolResult<Ev>(res).status).toBe('unknown');
    await h.close();
  });

  it('is annotated idempotent and not destructive', async () => {
    const { client } = stub(invite());
    const h = await harness(client);
    const tool = (await h.client.listTools()).tools.find((t) => t.name === 'outlook_respond_to_invite');
    expect(tool?.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    });
    await h.close();
  });
});
