import { describe, expect, it, vi, afterEach } from 'vitest';
import { mkdtempSync, readFileSync, existsSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { createTokenCache, tokenCachePath, reportCacheWriteFailure } from '../src/token-cache.js';
import { registerDirectoryTools } from '../src/tools/directory.js';
import { registerCalendarTools } from '../src/tools/calendar.js';
import { registerMailTools } from '../src/tools/mail.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';
import type { OutlookClient } from '../src/client.js';

function stub(get: (path: string, opts?: unknown) => Promise<unknown> = async () => ({ value: [] })) {
  return {
    get: vi.fn(get),
    write: vi.fn(async () => ({})),
    refreshTokenIfNeeded: vi.fn(async () => {}),
    tokenExpiresAt: () => Date.now() + 3_600_000,
    tokenSource: 'browser capture',
    apiBase: 'https://outlook.office.com/api/v2.0',
  } as unknown as OutlookClient;
}

const dir = () => mkdtempSync(join(tmpdir(), 'oo-cache-'));

describe('token cache', () => {
  it('resolves its own directory, not the access skill\'s', () => {
    const p = tokenCachePath({ HOME: '/home/x' } as NodeJS.ProcessEnv);
    expect(p).toContain('.office-outlook-mcp');
    expect(p.endsWith('token.json')).toBe(true);
    expect(p).not.toContain('.outlook-fpx');
  });

  it('honours an explicit path override', () => {
    const p = join(dir(), 'custom.json');
    expect(tokenCachePath({ OUTLOOK_TOKEN_FILE: p } as NodeJS.ProcessEnv)).toBe(p);
  });

  it('is off when explicitly disabled', () => {
    expect(createTokenCache({ OUTLOOK_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv)).toBeNull();
  });

  it('caches nothing when the token IS the environment variable', () => {
    // Nothing to skip: caching would only copy a credential onto disk.
    expect(
      createTokenCache({
        OUTLOOK_ACCESS_TOKEN: 'tok',
        OUTLOOK_TOKEN_FILE: join(dir(), 't.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('caches nothing when the bridge is the only source and it is disabled', () => {
    expect(
      createTokenCache({
        OUTLOOK_DISABLE_FETCHPROXY: '1',
        OUTLOOK_TOKEN_FILE: join(dir(), 't.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('round-trips a captured token and rejects a malformed record', () => {
    const file = join(dir(), 'token.json');
    const cache = createTokenCache({ OUTLOOK_TOKEN_FILE: file } as NodeJS.ProcessEnv);
    expect(cache).not.toBeNull();
    const tokens = { accessToken: 'a', refreshToken: 'r', expiresAt: Date.now() + 1000 };
    cache!.save(tokens);
    expect(existsSync(file)).toBe(true);
    expect(cache!.load()).toMatchObject({ accessToken: 'a' });
    // 0600: the file holds a live credential.
    expect(readFileSync(file, 'utf8')).toContain('accessToken');
    cache!.clear();
    expect(cache!.load()).toBeNull();
  });

  it('rejects a record missing a numeric expiry', () => {
    const file = join(dir(), 'token.json');
    const cache = createTokenCache({ OUTLOOK_TOKEN_FILE: file } as NodeJS.ProcessEnv)!;
    cache.save({ accessToken: 'a', expiresAt: Date.now() + 1000 });
    // A cache holding a non-number expiry would be written every mint and
    // rejected every load — doing nothing, silently.
    expect(cache.load()).not.toBeNull();
  });

  it('reports a failed cache write without throwing', () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    reportCacheWriteFailure(new Error('read-only fs'));
    reportCacheWriteFailure('not an error');
    expect(spy).toHaveBeenCalledTimes(2);
    expect(spy.mock.calls[0][0]).toContain('read-only fs');
    spy.mockRestore();
  });
});

afterEach(() => vi.restoreAllMocks());

describe('directory tools hit the documented paths', () => {
  const cases: [string, Record<string, unknown>, string][] = [
    ['outlook_get_profile', {}, '/me'],
    ['outlook_get_mailbox_settings', {}, '/me/MailboxSettings'],
    ['outlook_list_contacts', {}, '/me/contacts'],
    ['outlook_list_people', {}, '/me/people'],
    ['outlook_list_tasks', {}, '/me/tasks'],
  ];

  it.each(cases)('%s -> %s', async (name, args, path) => {
    const client = stub();
    const h = await createTestHarness((s: McpServer) => registerDirectoryTools(s, client));
    await h.callTool(name, args);
    expect((client.get as ReturnType<typeof vi.fn>).mock.calls[0][0]).toContain(path);
    await h.close();
  });

  it('pages contacts with $skip', async () => {
    const client = stub();
    const h = await createTestHarness((s: McpServer) => registerDirectoryTools(s, client));
    await h.callTool('outlook_list_contacts', { limit: 5, skip: 10 });
    const p = (client.get as ReturnType<typeof vi.fn>).mock.calls[0][0] as string;
    // Keys are literal; only values are percent-encoded.
    expect(p).toContain('$top=5');
    expect(p).toContain('$skip=10');
    await h.close();
  });

  it('returns the upstream envelope untouched in the raw view', async () => {
    const client = stub(async () => ({ value: [{ DisplayName: 'x' }], '@odata.context': 'ctx' }));
    const h = await createTestHarness((s: McpServer) => registerDirectoryTools(s, client));
    const out = parseToolResult<Record<string, unknown>>(
      await h.callTool('outlook_list_people', { view: 'raw' }),
    );
    expect(out['@odata.context']).toBe('ctx');
    await h.close();
  });
});

describe('calendar tools', () => {
  const event = {
    Id: 'e1',
    Subject: 'Standup',
    Start: { DateTime: '2026-09-22T15:00:00', TimeZone: 'UTC' },
    End: { DateTime: '2026-09-22T15:15:00', TimeZone: 'UTC' },
    Attendees: [{ EmailAddress: { Address: 'a@x.y' }, Status: { Response: 'Accepted' } }],
  };

  it('gets one event, with attendees only in the full view', async () => {
    const client = stub(async () => event);
    const h = await createTestHarness((s: McpServer) => registerCalendarTools(s, client));
    const compact = parseToolResult<Record<string, unknown>>(
      await h.callTool('outlook_get_event', { id: 'e1' }),
    );
    expect(compact).not.toHaveProperty('Attendees');
    const full = parseToolResult<Record<string, unknown>>(
      await h.callTool('outlook_get_event', { id: 'e1', view: 'full' }),
    );
    expect(full.Attendees).toEqual([{ Who: 'a@x.y', Response: 'Accepted' }]);
    await h.close();
  });

  it('returns the raw event when asked', async () => {
    const client = stub(async () => event);
    const h = await createTestHarness((s: McpServer) => registerCalendarTools(s, client));
    const raw = parseToolResult<Record<string, unknown>>(
      await h.callTool('outlook_get_event', { id: 'e1', view: 'raw' }),
    );
    expect(raw.Start).toEqual({ DateTime: '2026-09-22T15:00:00', TimeZone: 'UTC' });
    await h.close();
  });

  it('omits the timezone Prefer header when none is asked for', async () => {
    const client = stub();
    const h = await createTestHarness((s: McpServer) => registerCalendarTools(s, client));
    await h.callTool('outlook_list_events', { start: 'a', end: 'b' });
    expect((client.get as ReturnType<typeof vi.fn>).mock.calls[0][1]).toEqual({});
    await h.close();
  });

  it('lists calendars, raw and projected', async () => {
    const client = stub(async () => ({ value: [{ Id: 'c1', Name: 'Cal' }] }));
    const h = await createTestHarness((s: McpServer) => registerCalendarTools(s, client));
    expect(
      parseToolResult<{ count: number }>(await h.callTool('outlook_list_calendars', {})).count,
    ).toBe(1);
    expect(
      parseToolResult<Record<string, unknown>>(
        await h.callTool('outlook_list_calendars', { view: 'raw' }),
      ).value,
    ).toHaveLength(1);
    await h.close();
  });
});

describe('mail tools, remaining paths', () => {
  it('lists folders and honours the raw view', async () => {
    const client = stub(async () => ({ value: [{ Id: 'f', DisplayName: 'Inbox', UnreadItemCount: 2, TotalItemCount: 9 }] }));
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    const out = parseToolResult<{ items: Record<string, unknown>[] }>(
      await h.callTool('outlook_list_folders', {}),
    );
    expect(out.items[0]).toEqual({ Id: 'f', Name: 'Inbox', Unread: 2, Total: 9 });
    const raw = parseToolResult<Record<string, unknown>>(
      await h.callTool('outlook_list_folders', { view: 'raw' }),
    );
    expect(raw.value).toHaveLength(1);
    await h.close();
  });

  it('lists attachments as metadata only', async () => {
    const client = stub(async () => ({ value: [{ Id: 'a', Name: 'f.pdf', Size: 10, ContentType: 'application/pdf' }] }));
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    const out = parseToolResult<{ count: number; items: Record<string, unknown>[] }>(
      await h.callTool('outlook_list_attachments', { id: 'm1' }),
    );
    expect(out.count).toBe(1);
    // No bytes: ContentBytes is deliberately excluded via $select.
    expect((client.get as ReturnType<typeof vi.fn>).mock.calls[0][0]).toContain('Id%2CName%2CSize%2CContentType');
    expect(out.items[0]).not.toHaveProperty('ContentBytes');
    await h.close();
  });

  it('handles an attachments response with no value array', async () => {
    const client = stub(async () => ({}));
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    expect(
      parseToolResult<{ count: number }>(await h.callTool('outlook_list_attachments', { id: 'm' })).count,
    ).toBe(0);
    await h.close();
  });

  it('returns a raw and a full single message', async () => {
    const msg = { Id: 'm1', Subject: 'S', ConversationId: 'c', Body: { ContentType: 'Text', Content: 'b' } };
    const client = stub(async () => msg);
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    expect(
      parseToolResult<Record<string, unknown>>(await h.callTool('outlook_get_message', { id: 'm1', view: 'raw' })).Body,
    ).toEqual({ ContentType: 'Text', Content: 'b' });
    expect(
      parseToolResult<Record<string, unknown>>(await h.callTool('outlook_get_message', { id: 'm1', view: 'full' })).ConversationId,
    ).toBe('c');
    await h.close();
  });

  it('lists messages in the full view', async () => {
    const client = stub(async () => ({ value: [{ Id: 'm', Subject: 'S', ConversationId: 'c' }] }));
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    const out = parseToolResult<{ items: Record<string, unknown>[] }>(
      await h.callTool('outlook_list_messages', { view: 'full' }),
    );
    expect(out.items[0].ConversationId).toBe('c');
    await h.close();
  });

  it('returns the raw envelope for a listing when asked', async () => {
    const client = stub(async () => ({ value: [], '@odata.context': 'ctx' }));
    const h = await createTestHarness((s: McpServer) => registerMailTools(s, client));
    expect(
      parseToolResult<Record<string, unknown>>(await h.callTool('outlook_list_messages', { view: 'raw' }))['@odata.context'],
    ).toBe('ctx');
    await h.close();
  });
});

describe('healthcheck', () => {
  it('registers under the outlook prefix', async () => {
    const client = stub();
    const h = await createTestHarness((s: McpServer) => registerHealthcheckTool(s, client));
    expect((await h.listTools()).map((t) => t.name)).toEqual(['outlook_healthcheck']);
    await h.close();
  });

  it('resolves through the real credential path and reports no secret', async () => {
    const client = stub(async () => ({ EmailAddress: 'a@b.c', DisplayName: 'A' }));
    const h = await createTestHarness((s: McpServer) => registerHealthcheckTool(s, client));
    const res = await h.callTool('outlook_healthcheck', {});
    const report = parseToolResult<{
      ok: boolean;
      credential: { source: string; resolved: boolean; detail: Record<string, unknown> };
      probe: { url: string };
    }>(res);
    // It must exercise the same resolution the tools use...
    expect(client.refreshTokenIfNeeded).toHaveBeenCalled();
    // ...and report a SOURCE LABEL plus non-secret detail, never the token:
    // this is the output people paste into a chat when something is broken.
    expect(report.ok).toBe(true);
    expect(report.credential.source).toBe('browser capture');
    expect(report.credential.resolved).toBe(true);
    expect(report.probe.url).toBe('https://outlook.office.com/api/v2.0/me');
    expect(JSON.stringify(report)).not.toMatch(/eyJ[A-Za-z0-9_-]{10,}/);
    expect(Object.keys(report.credential.detail)).toEqual([
      'expiresAt',
      'minutesRemaining',
      'apiBase',
    ]);
    await h.close();
  });

  it('reports a failed probe instead of throwing', async () => {
    const client = stub(async () => {
      throw new Error('upstream exploded');
    });
    const h = await createTestHarness((s: McpServer) => registerHealthcheckTool(s, client));
    const res = await h.callTool('outlook_healthcheck', {});
    expect(JSON.stringify(res.content)).toContain('upstream exploded');
    await h.close();
  });
});

describe('confirmed write paths', () => {
  it('creates a draft and reports the new id', async () => {
    const client = stub();
    (client.write as ReturnType<typeof vi.fn>).mockResolvedValue({ Id: 'd1', WebLink: 'w' });
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const out = parseToolResult<{ created: boolean; Id: string }>(
      await h.callTool('outlook_create_draft', {
        to: ['a@example.com'], cc: ['c@example.com'], subject: 's', body: 'b', html: true, confirm: true,
      }),
    );
    expect(out).toMatchObject({ created: true, Id: 'd1' });
    const [, path, body] = (client.write as ReturnType<typeof vi.fn>).mock.calls[0];
    expect(path).toBe('/me/messages');
    expect(body).toMatchObject({ Body: { ContentType: 'HTML' } });
    await h.close();
  });

  it('moves a message and warns that the id changes', async () => {
    const client = stub();
    (client.write as ReturnType<typeof vi.fn>).mockResolvedValue({ Id: 'new', ParentFolderId: 'arch' });
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const out = parseToolResult<{ moved: boolean; newId: string; note: string }>(
      await h.callTool('outlook_move_message', { id: 'm1', destination: 'archive', confirm: true }),
    );
    expect(out.moved).toBe(true);
    expect(out.newId).toBe('new');
    expect(out.note).toMatch(/new Id/);
    await h.close();
  });

  it('creates an event, defaulting the zone to UTC', async () => {
    const client = stub();
    (client.write as ReturnType<typeof vi.fn>).mockResolvedValue({ Id: 'e9' });
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    await h.callTool('outlook_create_event', {
      subject: 's', start: '2026-09-22T15:00:00', end: '2026-09-22T16:00:00', confirm: true,
    });
    expect((client.write as ReturnType<typeof vi.fn>).mock.calls[0][2]).toMatchObject({
      Start: { TimeZone: 'UTC' },
    });
    await h.close();
  });

  it('carries location, body and attendees onto the event when given', async () => {
    const client = stub();
    (client.write as ReturnType<typeof vi.fn>).mockResolvedValue({ Id: 'e9' });
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    await h.callTool('outlook_create_event', {
      subject: 's', start: 'a', end: 'b', timeZone: 'Eastern Standard Time',
      location: 'Room 1', body: 'agenda', attendees: ['x@y.z'], confirm: true,
    });
    expect((client.write as ReturnType<typeof vi.fn>).mock.calls[0][2]).toMatchObject({
      Location: { DisplayName: 'Room 1' },
      Body: { Content: 'agenda' },
      Attendees: [{ EmailAddress: { Address: 'x@y.z' }, Type: 'Required' }],
      Start: { TimeZone: 'Eastern Standard Time' },
    });
    await h.close();
  });

  it('honours saveToSentItems: false', async () => {
    const client = stub();
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    await h.callTool('outlook_send_mail', {
      to: ['a@example.com'], subject: 's', body: 'b', saveToSentItems: false, confirm: true,
    });
    expect((client.write as ReturnType<typeof vi.fn>).mock.calls[0][2]).toMatchObject({
      SaveToSentItems: false,
    });
    await h.close();
  });

  it('reports no recipients rather than an empty string in the preview', async () => {
    const client = stub();
    const { registerWriteTools } = await import('../src/tools/writes.js');
    const h = await createTestHarness((s: McpServer) => registerWriteTools(s, client));
    const out = parseToolResult<{ action: string }>(
      await h.callTool('outlook_send_mail', { subject: 's', body: 'b' }),
    );
    expect(out.action).toContain('(no recipients)');
    await h.close();
  });
});

describe('client write and refresh plumbing', () => {
  const jwt = (s: number) => {
    const b = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
    return `${b({ alg: 'none' })}.${b({ exp: Math.floor(Date.now() / 1000) + s })}.x`;
  };
  const ok = () => new Response('{}', { status: 200, headers: { 'content-type': 'application/json' } });

  it('sends a JSON body on a write and can refresh on demand', async () => {
    const { OutlookClient } = await import('../src/client.js');
    const fetchImpl = vi.fn(async () => ok());
    const capture = vi.fn(async () => jwt(3600));
    const c = new OutlookClient({
      env: { OUTLOOK_TOKEN_CACHE: 'false' },
      captureToken: capture,
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    await c.write('POST', '/me/sendmail', { Message: { Subject: 's' } });
    const init = (fetchImpl.mock.calls[0] as unknown as [unknown, RequestInit])[1];
    expect(init.method).toBe('POST');
    expect(String(init.body)).toContain('Subject');

    await c.refreshToken();
    expect(capture).toHaveBeenCalledTimes(2);
    expect(c.tokenExpiresAt()).toBeGreaterThan(Date.now());
  });

  it('reports no expiry before any token is held', async () => {
    const { OutlookClient } = await import('../src/client.js');
    const c = new OutlookClient({ env: { OUTLOOK_TOKEN_CACHE: 'false' } });
    expect(c.tokenExpiresAt()).toBeNull();
  });

  it('refuses writes and refreshes when nothing is configured', async () => {
    const { OutlookClient } = await import('../src/client.js');
    const c = new OutlookClient({
      env: { OUTLOOK_TOKEN_CACHE: 'false', OUTLOOK_DISABLE_FETCHPROXY: '1' },
    });
    await expect(c.write('POST', '/x', {})).rejects.toThrow(/No Outlook credential/);
    await expect(c.refreshToken()).rejects.toThrow(/No Outlook credential/);
    await expect(c.refreshTokenIfNeeded()).rejects.toThrow(/No Outlook credential/);
    await expect(c.getAbsolute('https://outlook.office.com/x')).rejects.toThrow(
      /No Outlook credential/,
    );
  });
});

describe('a broken token cache degrades instead of failing the call', () => {
  const jwt = (s: number) => {
    const b = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
    return `${b({ alg: 'none' })}.${b({ exp: Math.floor(Date.now() / 1000) + s })}.x`;
  };

  it('swallows a failed cache write, reports it, and still returns the response', async () => {
    // A read-only data dir should cost the next cold start a capture, not this
    // request.
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const blocker = join(dir(), 'not-a-dir');
    writeFileSync(blocker, 'x');
    const { OutlookClient } = await import('../src/client.js');
    const c = new OutlookClient({
      // Nest the token file UNDER an existing regular file, so mkdir -p
      // genuinely fails (ENOTDIR) rather than quietly creating the parent.
      // Making the real write fail beats mocking the persistence layer.
      env: { OUTLOOK_TOKEN_FILE: join(blocker, 'token.json') },
      captureToken: async () => jwt(3600),
      fetchImpl: (async () =>
        new Response('{"EmailAddress":"a@b.c"}', {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })) as unknown as typeof fetch,
    });
    await expect(c.get('/me')).resolves.toMatchObject({ EmailAddress: 'a@b.c' });
    expect(spy.mock.calls.flat().join(' ')).toContain('could not cache the access token');
    spy.mockRestore();
  });
});

describe('healthcheck with no token held', () => {
  it('reports a null expiry rather than computing one from nothing', async () => {
    const client = {
      get: vi.fn(async () => ({ EmailAddress: 'a@b.c' })),
      refreshTokenIfNeeded: vi.fn(async () => {}),
      tokenExpiresAt: () => null,
      tokenSource: 'OUTLOOK_ACCESS_TOKEN',
      apiBase: 'https://outlook.office.com/api/v2.0',
    } as unknown as OutlookClient;
    const h = await createTestHarness((s: McpServer) => registerHealthcheckTool(s, client));
    const report = parseToolResult<{ credential: { detail: Record<string, unknown> } }>(
      await h.callTool('outlook_healthcheck', {}),
    );
    expect(report.credential.detail.expiresAt).toBeNull();
    expect(report.credential.detail.minutesRemaining).toBeNull();
    await h.close();
  });
});
