import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerMailTools } from '../src/tools/mail.js';
import { registerDraftTools } from '../src/tools/drafts.js';
import type { OutlookClient } from '../src/client.js';

/**
 * Rich bodies: an agent writes Markdown (cheaper than HTML, nothing to get
 * wrong) and reads Markdown or sanitised HTML instead of raw mail HTML.
 * Fixtures are synthetic.
 */

type Rec = Record<string, unknown>;
const addr = (Address: string) => ({ EmailAddress: { Address } });

/** A mail body as a sender's client writes it: hidden preheader, styles, a layout table, a list and a link. */
const MAIL_HTML =
  '<html><head><style>p{color:red}</style></head><body>' +
  '<div style="display:none;max-height:0">Preheader you never see</div>' +
  '<table width="100%"><tr><td><p style="font:14px x">Hi <b>Chris</b>,</p></td></tr>' +
  '<tr><td><ol><li>one</li><li>two</li></ol><p>See <a href="https://example.test/doc" style="color:blue">the doc</a>.</p></td></tr></table>' +
  '</body></html>';

function writeStub(draftBody: { ContentType: string; Content: string } = { ContentType: 'HTML', Content: '<html><body><div>quoted original</div></body></html>' }) {
  const calls: { method: string; path: string; body?: unknown }[] = [];
  const get = vi.fn(async () => ({
    Id: 'm-1',
    Subject: 'Budget review',
    From: addr('alice@example.test'),
    ToRecipients: [addr('me@example.test')],
    CcRecipients: [],
    ReplyTo: [],
  }));
  const write = vi.fn(async (method: string, path: string, body?: unknown) => {
    calls.push({ method, path, body });
    if (/\/create(reply|replyall|forward)$/.test(path)) return { Id: 'd-1', WebLink: 'https://x.test/d-1', Body: draftBody };
    if (method === 'PATCH') return { Id: 'd-1', ...(body as Rec) };
    if (path === '/me/messages') return { Id: 'new-1' };
    return undefined;
  });
  return { client: { get, write } as unknown as OutlookClient, calls };
}

const writes = (client: OutlookClient) => createTestHarness((s: McpServer) => registerWriteTools(s, client));

async function confirm(h: TestHarness, tool: string, args: Rec) {
  const first = parseToolResult<{ confirmToken?: string; status?: string }>(await h.callTool(tool, args));
  expect(first.status).toBe('confirmation-required');
  return h.callTool(tool, { ...args, confirmToken: first.confirmToken });
}

describe('writing Markdown', () => {
  it('outlook_send_message sends a Markdown body as formatted HTML', async () => {
    const { client, calls } = writeStub();
    const h = await writes(client);
    const res = await confirm(h, 'outlook_send_message', {
      to: ['bob@example.test'],
      subject: 'Plan',
      body: '## Plan\n\n1. **ship** it\n2. see [doc](https://example.test/d)\n\n> quoted <b>raw</b>',
      format: 'markdown',
    });
    expect(res.isError).toBeFalsy();
    const body = (calls[0].body as { Message: { Body: { ContentType: string; Content: string } } }).Message.Body;
    expect(body).toEqual({
      ContentType: 'HTML',
      Content:
        '<h2>Plan</h2><ol><li><strong>ship</strong> it</li><li>see <a href="https://example.test/d">doc</a></li></ol>' +
        '<blockquote><p>quoted &lt;b&gt;raw&lt;/b&gt;</p></blockquote>',
    });
    await h.close();
  });

  it('keeps plain text the default, and html:true as an alias for format "html"', async () => {
    const { client, calls } = writeStub();
    const h = await writes(client);
    await confirm(h, 'outlook_send_message', { to: ['b@example.test'], subject: 's', body: '**not bold**' });
    await confirm(h, 'outlook_send_message', { to: ['b@example.test'], subject: 's', body: '<i>x</i>', html: true });
    await confirm(h, 'outlook_send_message', { to: ['b@example.test'], subject: 's', body: '<i>x</i>', format: 'html' });
    const bodies = calls.map((c) => (c.body as { Message: { Body: unknown } }).Message.Body);
    expect(bodies).toEqual([
      { ContentType: 'Text', Content: '**not bold**' },
      { ContentType: 'HTML', Content: '<i>x</i>' },
      { ContentType: 'HTML', Content: '<i>x</i>' },
    ]);
    await h.close();
  });

  it('refuses html:true with a different format', async () => {
    const { client } = writeStub();
    const h = await writes(client);
    const res = await h.callTool('outlook_send_message', { to: ['b@example.test'], subject: 's', body: 'x', html: true, format: 'markdown' });
    expect(res.isError).toBe(true);
    expect(JSON.stringify(res)).toMatch(/html.*format/i);
    await h.close();
  });

  it('outlook_create_draft composes a Markdown body as HTML', async () => {
    const { client, calls } = writeStub();
    const h = await writes(client);
    await h.callTool('outlook_create_draft', { to: ['b@example.test'], subject: 's', body: '- a\n- *b*', format: 'markdown' });
    expect((calls[0].body as { Body: unknown }).Body).toEqual({ ContentType: 'HTML', Content: '<ul><li>a</li><li><em>b</em></li></ul>' });
    await h.close();
  });

  it('outlook_reply with a Markdown comment drafts, writes the formatted comment above the quote, and sends', async () => {
    const { client, calls } = writeStub();
    const h = await writes(client);
    const res = await confirm(h, 'outlook_reply', { messageId: 'm-1', mode: 'reply', comment: 'Thanks — **approved**.\n\n- next: ship', format: 'markdown' });
    expect(res.isError).toBeFalsy();
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual([
      'POST /me/messages/m-1/createreply',
      'PATCH /me/messages/d-1',
      'POST /me/messages/d-1/send',
    ]);
    const patch = calls[1].body as { Body: { ContentType: string; Content: string } };
    expect(patch.Body.ContentType).toBe('HTML');
    expect(patch.Body.Content).toBe(
      '<html><body><p>Thanks — <strong>approved</strong>.</p><ul><li>next: ship</li></ul><br><div>quoted original</div></body></html>',
    );
    expect(parseToolResult<Rec>(res)).toMatchObject({ sent: true, mode: 'reply' });
    await h.close();
  });

  it('turns a plain-text reply draft into HTML so a Markdown comment can be formatted, keeping the quote as text', async () => {
    const { client, calls } = writeStub({ ContentType: 'Text', Content: '> From: Alice\n> a < b & c' });
    const h = await writes(client);
    await h.callTool('outlook_reply', { messageId: 'm-1', mode: 'reply', comment: '**ok**', format: 'markdown', draftOnly: true });
    const patch = calls[1].body as { Body: { ContentType: string; Content: string } };
    expect(patch.Body).toEqual({
      ContentType: 'HTML',
      Content: '<p><strong>ok</strong></p><br><div>&gt; From: Alice<br>&gt; a &lt; b &amp; c</div>',
    });
    await h.close();
  });

  it('keeps a plain-text reply a single send, and binds the token to the format', async () => {
    const { client, calls } = writeStub();
    const h = await writes(client);
    const args = { messageId: 'm-1', mode: 'reply', comment: '**ok**' };
    const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_reply', args));
    // A token from the plain-text preview cannot send the Markdown version.
    const swapped = parseToolResult<{ status?: string }>(
      await h.callTool('outlook_reply', { ...args, format: 'markdown', confirmToken: first.confirmToken }),
    );
    expect(swapped.status).not.toBe(undefined);
    expect(calls).toHaveLength(0);
    await h.callTool('outlook_reply', { ...args, confirmToken: first.confirmToken });
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual(['POST /me/messages/m-1/reply']);
    expect(calls[0].body).toEqual({ Comment: '**ok**' });
    await h.close();
  });
});

describe('outlook_update_draft with Markdown', () => {
  it('replaces the body with formatted HTML', async () => {
    let current: Rec = {
      Id: 'd1',
      IsDraft: true,
      Subject: 's',
      Body: { ContentType: 'Text', Content: 'old' },
      ToRecipients: [],
      CcRecipients: [],
      BccRecipients: [],
    };
    const get = vi.fn(async () => current);
    const write = vi.fn(async (_m: string, _p: string, body?: Rec) => {
      const b = body?.Body as { ContentType: string; Content: string } | undefined;
      current = { ...current, ...body, ...(b ? { Body: { ContentType: b.ContentType, Content: `<html><body>${b.Content}</body></html>` } } : {}) };
      return current;
    });
    const client = { get, write } as unknown as OutlookClient;
    const h = await createTestHarness((s: McpServer) => registerDraftTools(s, client, { sendCheckDelayMs: 0 }));
    const first = parseToolResult<{ confirmToken?: string }>(await h.callTool('outlook_update_draft', { id: 'd1', body: '**new**', format: 'markdown' }));
    const res = parseToolResult<Rec>(await h.callTool('outlook_update_draft', { id: 'd1', body: '**new**', format: 'markdown', confirmToken: first.confirmToken }));
    expect(res.updated).toBe(true);
    expect(write.mock.calls[0][2]).toMatchObject({ Body: { ContentType: 'HTML', Content: '<p><strong>new</strong></p>' } });
    const noBody = await h.callTool('outlook_update_draft', { id: 'd1', subject: 'x', format: 'markdown' });
    expect(noBody.isError).toBe(true);
    await h.close();
  });
});

describe('reading Markdown or sanitised HTML', () => {
  function readStub(body: { ContentType: string; Content: string } = { ContentType: 'HTML', Content: MAIL_HTML }) {
    const get = vi.fn(async (path: string, opts?: { text?: boolean; query?: Rec }) => {
      if (path.includes('/messages/') ) {
        return { Id: 'm1', Subject: 'Hi', From: addr('a@example.test'), Body: opts?.text ? { ContentType: 'Text', Content: 'plain' } : body };
      }
      // the unread list
      return { value: [{ Id: 'm1', Subject: 'Hi', From: addr('a@example.test'), ReceivedDateTime: '2026-11-11T10:00:00Z', IsRead: false }] };
    });
    return { client: { get, write: vi.fn() } as unknown as OutlookClient, get };
  }
  const mail = (client: OutlookClient) => createTestHarness((s: McpServer) => registerMailTools(s, client));

  it('format "markdown" fetches the HTML and returns compact Markdown, hidden preheader dropped', async () => {
    const { client, get } = readStub();
    const h = await mail(client);
    const res = parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1', format: 'markdown' }));
    expect(get.mock.calls[0][1]).toMatchObject({ text: false });
    expect(res.Body).toBe('Hi **Chris**,\n\n1. one\n2. two\n\nSee [the doc](https://example.test/doc).');
    await h.close();
  });

  it('format "html" returns sanitised HTML; html:true keeps the raw HTML it always returned; view "raw" stays untouched', async () => {
    const { client } = readStub();
    const h = await mail(client);
    const clean =
      '<table><tr><td><p>Hi <b>Chris</b>,</p></td></tr><tr><td><ol><li>one</li><li>two</li></ol>' +
      '<p>See <a href="https://example.test/doc">the doc</a>.</p></td></tr></table>';
    expect(parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1', format: 'html' })).Body).toBe(clean);
    // html:true predates `format` and returned the sender's HTML as Outlook stores it; it still does.
    expect(parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1', html: true })).Body).toBe(MAIL_HTML);
    expect(parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1', html: true, format: 'html' })).Body).toBe(clean);
    const raw = parseToolResult<{ Body: { Content: string } }>(await h.callTool('outlook_get_message', { id: 'm1', format: 'markdown', view: 'raw' }));
    expect(raw.Body.Content).toBe(MAIL_HTML);
    await h.close();
  });

  it('leaves a plain-text mail as it is, and the default as Outlook\'s plain text', async () => {
    const { client } = readStub({ ContentType: 'Text', Content: 'a < b *as typed*' });
    const h = await mail(client);
    expect(parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1', format: 'markdown' })).Body).toBe('a < b *as typed*');
    expect(parseToolResult<Rec>(await h.callTool('outlook_get_message', { id: 'm1' })).Body).toBe('plain');
    await h.close();
  });

  it('outlook_list_unread_messages reads bodies as Markdown with bodyFormat "markdown"', async () => {
    const { client, get } = readStub();
    const h = await mail(client);
    const res = parseToolResult<{ items: Rec[] }>(await h.callTool('outlook_list_unread_messages', { bodyFormat: 'markdown' }));
    expect(res.items[0].body).toBe('Hi **Chris**,\n\n1. one\n2. two\n\nSee [the doc](https://example.test/doc).');
    const one = get.mock.calls.find((c) => String(c[0]).startsWith('/me/messages/'))!;
    expect(one[1]).toMatchObject({ text: false });
    await h.close();
  });
});
