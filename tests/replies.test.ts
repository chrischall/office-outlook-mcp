import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { buildQueryString, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import type { OutlookClient } from '../src/client.js';

/**
 * `outlook_reply` — reply / reply-all / forward a message, sent at once
 * (behind the confirm gate) or left as a draft. Fixtures are synthetic
 * (example.test addresses, placeholder subjects and text).
 */

type Rec = Record<string, unknown>;

const wire = (path: string, opts?: { query?: Record<string, unknown> }) =>
  `${path}${buildQueryString((opts?.query ?? {}) as Record<string, string | number | undefined>)}`;

const addr = (Address: string) => ({ EmailAddress: { Address } });

function original(over: Rec = {}): Rec {
  return {
    Id: 'm-1',
    Subject: 'Budget review',
    From: addr('alice@example.test'),
    ToRecipients: [addr('me@example.test'), addr('bob@example.test')],
    CcRecipients: [addr('carol@example.test')],
    ReplyTo: [],
    ...over,
  };
}

/** A draft as createreply/createforward returns it: quoted original in the body. */
function draft(over: Rec = {}): Rec {
  return {
    Id: 'd-1',
    WebLink: 'https://outlook.example.test/d-1',
    Subject: 'RE: Budget review',
    Body: { ContentType: 'HTML', Content: '<html><body><div>quoted original</div></body></html>' },
    ...over,
  };
}

function stub(opts: { msg?: Rec; draft?: Rec; writeError?: Error } = {}) {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  const get = vi.fn(async (p: string, o?: { query?: Record<string, unknown> }) => {
    calls.push({ method: 'GET', path: wire(p, o) });
    return opts.msg ?? original();
  });
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    if (opts.writeError) throw opts.writeError;
    if (/\/create(reply|replyall|forward)$/.test(path)) return opts.draft ?? draft();
    if (method === 'PATCH') return { ...(opts.draft ?? draft()), ...(body as Rec) };
    return undefined;
  });
  const client = { get, write } as unknown as OutlookClient;
  const writes = () => calls.filter((c) => c.method !== 'GET');
  return { client, calls, writes };
}

async function harness(client: OutlookClient) {
  return createTestHarness((s: McpServer) => registerWriteTools(s, client));
}

async function confirmed(h: TestHarness, args: Rec) {
  const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_reply', args));
  return h.callTool('outlook_reply', { ...args, confirmToken: first.confirmToken });
}

describe('outlook_reply — sending', () => {
  it('previews a reply to the sender with subject and comment, and writes nothing without a token', async () => {
    const { client, writes } = stub();
    const h = await harness(client);
    const res = parseToolResult<{ status?: string }>(
      await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'reply', comment: 'Thanks, looks good' }),
    );
    expect(res.status).toBe('confirmation-required');
    const text = JSON.stringify(res);
    expect(text).toMatch(/RE: Budget review/);
    expect(text).toMatch(/alice@example\.test/);
    expect(text).toMatch(/Thanks, looks good/);
    // A plain reply goes to the sender only.
    expect(text).not.toMatch(/bob@example\.test/);
    expect(text).not.toMatch(/carol@example\.test/);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('previews reply-all with the original To and Cc', async () => {
    const { client } = stub();
    const h = await harness(client);
    const text = JSON.stringify(
      parseToolResult(await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'replyAll', comment: 'ok' })),
    );
    expect(text).toMatch(/alice@example\.test/);
    expect(text).toMatch(/bob@example\.test/);
    expect(text).toMatch(/carol@example\.test/);
    await h.close();
  });

  it('replies to the Reply-To address when the original sets one', async () => {
    const { client } = stub({ msg: original({ ReplyTo: [addr('list@example.test')] }) });
    const h = await harness(client);
    const text = JSON.stringify(
      parseToolResult(await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'reply', comment: 'ok' })),
    );
    expect(text).toMatch(/list@example\.test/);
    await h.close();
  });

  it.each([
    ['reply', 'reply', undefined],
    ['replyAll', 'replyall', undefined],
    ['forward', 'forward', ['dave@example.test']],
  ])('%s posts { Comment } to /%s once confirmed', async (mode, verb, to) => {
    const { client, writes } = stub();
    const h = await harness(client);
    const res = await confirmed(h, { messageId: 'm-1', mode, comment: 'see below', ...(to ? { to } : {}) });
    expect(res.isError).toBeFalsy();
    expect(writes()).toEqual([
      {
        method: 'POST',
        path: `/me/messages/m-1/${verb}`,
        body: { Comment: 'see below', ...(to ? { ToRecipients: [addr('dave@example.test')] } : {}) },
      },
    ]);
    expect(parseToolResult<Rec>(res)).toMatchObject({ sent: true, mode });
    await h.close();
  });

  it('previews a forward with its new recipients and FW: subject', async () => {
    const { client } = stub();
    const h = await harness(client);
    const text = JSON.stringify(
      parseToolResult(
        await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'forward', comment: 'fyi', to: ['dave@example.test'] }),
      ),
    );
    expect(text).toMatch(/FW: Budget review/);
    expect(text).toMatch(/dave@example\.test/);
    await h.close();
  });

  it('refuses a forward with no recipients', async () => {
    const { client, writes } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'forward', comment: 'fyi' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/forward/i);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('refuses `to` on a reply, where Outlook picks the recipients', async () => {
    const { client, writes } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_reply', {
      messageId: 'm-1',
      mode: 'reply',
      comment: 'x',
      to: ['dave@example.test'],
    });
    expect(res.isError).toBe(true);
    expect(writes()).toHaveLength(0);
    await h.close();
  });

  it('reports an unconfirmed send as unknown and points at Sent Items', async () => {
    const { client } = stub({ writeError: new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await confirmed(h, { messageId: 'm-1', mode: 'reply', comment: 'x' });
    expect(res.isError).toBeFalsy();
    const body = parseToolResult<Rec>(res);
    expect(body.status).toBe('unknown');
    expect(String(body.warning)).toMatch(/Sent Items/);
    await h.close();
  });
});

describe('outlook_reply — draftOnly', () => {
  it('creates the reply draft and writes the comment above the quoted original, with no confirm step', async () => {
    const { client, writes } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_reply', {
      messageId: 'm-1',
      mode: 'replyAll',
      comment: 'Line one\nA & B <ok>',
      draftOnly: true,
    });
    expect(res.isError).toBeFalsy();
    const w = writes();
    expect(w).toHaveLength(2);
    expect(w[0]).toMatchObject({ method: 'POST', path: '/me/messages/m-1/createreplyall' });
    expect(w[1]).toMatchObject({ method: 'PATCH', path: '/me/messages/d-1' });
    const content = (w[1].body as { Body: { ContentType: string; Content: string } }).Body;
    expect(content.ContentType).toBe('HTML');
    // Escaped, line breaks kept, and placed inside <body> before the quote.
    expect(content.Content).toMatch(/<body><div>Line one<br>A &amp; B &lt;ok&gt;<\/div><br><div>quoted original/);
    expect(parseToolResult<Rec>(res)).toMatchObject({ drafted: true, draftId: 'd-1', mode: 'replyAll' });
    await h.close();
  });

  it('prepends to a plain-text draft body as text', async () => {
    const { client, writes } = stub({ draft: draft({ Body: { ContentType: 'Text', Content: '> quoted' } }) });
    const h = await harness(client);
    await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'reply', comment: 'hi', draftOnly: true });
    expect(writes()[1].body).toEqual({ Body: { ContentType: 'Text', Content: 'hi\n\n> quoted' } });
    await h.close();
  });

  it('sets the forward draft recipients', async () => {
    const { client, writes } = stub();
    const h = await harness(client);
    await h.callTool('outlook_reply', {
      messageId: 'm-1',
      mode: 'forward',
      comment: 'fyi',
      to: ['dave@example.test'],
      draftOnly: true,
    });
    const w = writes();
    expect(w[0]).toMatchObject({ method: 'POST', path: '/me/messages/m-1/createforward' });
    expect(w[1].body).toMatchObject({ ToRecipients: [addr('dave@example.test')] });
    await h.close();
  });
});

describe('outlook_reply — annotations', () => {
  it('is destructive: a sent reply cannot be unsent', async () => {
    const { client } = stub();
    const h = await harness(client);
    const tool = (await h.client.listTools()).tools.find((t) => t.name === 'outlook_reply');
    expect(tool?.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: true, openWorldHint: true });
    await h.close();
  });
});
