import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult, type TestHarness } from '@chrischall/mcp-utils/test';
import { ApiError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerDraftTools } from '../src/tools/drafts.js';
import type { OutlookClient } from '../src/client.js';

/**
 * Editing drafts. Shapes match what PATCH /me/messages/{id} did live on
 * 2026-11-10: it answers with the whole updated message; a Text body comes back
 * as sent; an HTML body comes back wrapped in <html><head>…<body>; `[]` clears a
 * recipient list; sent mail reads IsDraft false. Fixtures are synthetic.
 */

type Obj = Record<string, unknown>;
const addr = (a: string) => ({ EmailAddress: { Name: a, Address: a } });

function draft(over: Obj = {}): Obj {
  return {
    '@odata.id': 'x',
    Id: 'd1',
    IsDraft: true,
    Subject: 'Quarterly numbers',
    Body: { ContentType: 'Text', Content: 'Old text the user typed in Outlook' },
    ToRecipients: [addr('ann@example.test')],
    CcRecipients: [],
    BccRecipients: [],
    WebLink: 'https://outlook.example.test/d1',
    ...over,
  };
}

/** A mailbox held in memory; `persist: false` accepts writes but keeps the old message. */
function stub(opts: { messages?: Obj[]; persist?: boolean; writeError?: Error; emptyPatchReply?: boolean } = {}) {
  let messages = (opts.messages ?? [draft(), draft({ Id: 'sent1', IsDraft: false, Subject: 'Already sent' })]).map((m) => ({ ...m }));
  const find = (id: string) => {
    const m = messages.find((x) => x.Id === id);
    if (!m) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
    return m;
  };
  const idOf = (path: string) => decodeURIComponent((path.split('/')[3] ?? '').split('?')[0] ?? '');
  const get = vi.fn(async (path: string): Promise<unknown> => find(idOf(path)));
  const write = vi.fn(async (method: string, path: string, body?: Obj) => {
    if (opts.writeError) throw opts.writeError;
    if (method === 'POST' && path.endsWith('/send')) {
      // Outlook answers 202 with no body and files the sent copy under a new id.
      const sentDraft = find(idOf(path));
      if (opts.persist !== false) messages = messages.filter((m) => m.Id !== sentDraft.Id);
      return undefined;
    }
    if (method !== 'PATCH') throw new Error(`unexpected ${method} ${path}`);
    const current = find(idOf(path));
    const sent = { ...body } as Obj;
    const b = sent.Body as Obj | undefined;
    if (b?.ContentType === 'HTML') {
      sent.Body = { ContentType: 'HTML', Content: `<html><head>\r\n<meta charset="utf-8"></head><body>${String(b.Content)}</body></html>` };
    }
    const next = { ...current, ...sent };
    if (opts.persist !== false) messages = messages.map((m) => (m.Id === current.Id ? next : m));
    return opts.emptyPatchReply ? undefined : opts.persist === false ? current : next;
  });
  return { client: { get, write } as unknown as OutlookClient, get, write, now: (id = 'd1') => find(id) };
}

// No real waiting in tests: the send check polls with a zero delay.
const harness = (client: OutlookClient) =>
  createTestHarness((s: McpServer) => registerDraftTools(s, client, { sendCheckDelayMs: 0 }));

type Result = Obj & { status?: string; confirmToken?: string; warning?: string; preview?: Obj };

async function call(h: TestHarness, args: Obj): Promise<Result> {
  return parseToolResult<Result>(await h.callTool('outlook_update_draft', args));
}

async function confirmed(h: TestHarness, args: Obj): Promise<Result> {
  const first = await call(h, args);
  expect(first.status).toBe('confirmation-required');
  return call(h, { ...args, confirmToken: first.confirmToken });
}

const text = (res: { content?: unknown }) => (res.content as Array<{ text: string }>)[0]?.text ?? '';

describe('outlook_update_draft', () => {
  it('previews the draft as it is now and writes nothing without a token', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await call(h, { id: 'd1', body: 'New text' });
    expect(res.status).toBe('confirmation-required');
    const shown = JSON.stringify(res);
    expect(shown).toMatch(/Quarterly numbers/);
    expect(shown).toMatch(/Old text the user typed in Outlook/);
    expect(shown).toMatch(/replace/i);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('changes only what is passed, and verifies against what Outlook returns', async () => {
    const { client, write, now } = stub();
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', subject: 'Q3 numbers', body: 'Line one\nLine two' });
    expect(write.mock.calls).toEqual([
      ['PATCH', '/me/messages/d1', { Subject: 'Q3 numbers', Body: { ContentType: 'Text', Content: 'Line one\nLine two' } }],
    ]);
    expect(res).toMatchObject({ updated: true, Id: 'd1', Subject: 'Q3 numbers', To: ['ann@example.test'] });
    expect(res).not.toHaveProperty('@odata.id');
    expect(now().Subject).toBe('Q3 numbers');
    await h.close();
  });

  it('replaces each recipient list it is given, and [] clears one', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', to: [], cc: ['Bob@Example.test'], bcc: ['carol@example.test'] });
    expect(write.mock.calls[0]?.[2]).toEqual({
      ToRecipients: [],
      CcRecipients: [{ EmailAddress: { Address: 'Bob@Example.test' } }],
      BccRecipients: [{ EmailAddress: { Address: 'carol@example.test' } }],
    });
    expect(res).toMatchObject({ updated: true, To: [], Cc: ['Bob@Example.test'], Bcc: ['carol@example.test'] });
    await h.close();
  });

  it('sends an HTML body as HTML and accepts the wrapper Outlook adds', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', body: '<p>Hi <b>there</b></p>', html: true });
    expect(write.mock.calls[0]?.[2]).toEqual({ Body: { ContentType: 'HTML', Content: '<p>Hi <b>there</b></p>' } });
    expect(res).toMatchObject({ updated: true });
    expect(res.warning).toBeUndefined();
    await h.close();
  });

  it('re-reads the draft when the PATCH answers without a body', async () => {
    const { client, get } = stub({ emptyPatchReply: true });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', subject: 'Q3 numbers' });
    expect(res).toMatchObject({ updated: true, Subject: 'Q3 numbers' });
    expect(get.mock.calls.length).toBeGreaterThanOrEqual(2);
    await h.close();
  });

  it('warns, naming the fields, when Outlook accepts the edit but the draft did not change', async () => {
    const { client } = stub({ persist: false });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', subject: 'Q3 numbers', to: ['zed@example.test'] });
    expect(res).toMatchObject({ updated: false });
    expect(res.warning).toMatch(/Subject/);
    expect(res.warning).toMatch(/To/);
    await h.close();
  });

  it('refuses a message that is not a draft, before asking anything', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_draft', { id: 'sent1', subject: 'x' });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/not a draft/i);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('explains an id that matches no message', async () => {
    const { client } = stub();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_draft', { id: 'nope', subject: 'x' });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/No message matches "nope"/);
    await h.close();
  });

  it('passes a server error on the read straight through', async () => {
    const { client, get } = stub();
    get.mockRejectedValueOnce(new ApiError(500, 'Outlook 500'));
    const h = await harness(client);
    const res = await h.callTool('outlook_update_draft', { id: 'd1', subject: 'x' });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/500/);
    await h.close();
  });

  it.each([
    ['nothing to change', { id: 'd1' }],
    ['html without a body', { id: 'd1', subject: 'x', html: true }],
  ])('refuses %s', async (_label, args) => {
    const { client, write } = stub();
    const h = await harness(client);
    expect((await h.callTool('outlook_update_draft', args)).isError).toBe(true);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('reports an unknown outcome instead of inviting a blind retry', async () => {
    const { client } = stub({ writeError: new WriteOutcomeUnknownError('Outlook', 'PATCH', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', subject: 'x' });
    expect(res.status).toBe('unknown');
    await h.close();
  });

  it('says to check before retrying when the edit lands but cannot be re-read', async () => {
    const { client, get, write } = stub({ emptyPatchReply: true });
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    get.mockImplementation(async (p: string) => {
      if (write.mock.calls.length > 0) throw new ApiError(503, 'Outlook 503');
      return real(p);
    });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', subject: 'Q3' });
    expect(res).toMatchObject({ updated: true });
    expect(res.warning).toMatch(/before retrying/i);
    await h.close();
  });
});

describe('outlook_update_draft — addCc / removeCc', () => {
  const withCc = () => stub({ messages: [draft({ CcRecipients: [addr('naresh@example.test')] })] });

  it('adds a cc while keeping the existing one, and shows the resulting Cc in the preview', async () => {
    const { client, write, now } = withCc();
    const h = await harness(client);
    const first = await call(h, { id: 'd1', addCc: ['kiran@example.test'] });
    expect(first.status).toBe('confirmation-required');
    const shown = JSON.stringify(first);
    expect(shown).toMatch(/naresh@example\.test/);
    expect(shown).toMatch(/kiran@example\.test/);
    expect(write).not.toHaveBeenCalled();
    const res = await call(h, { id: 'd1', addCc: ['kiran@example.test'], confirmToken: first.confirmToken });
    expect(res).toMatchObject({ updated: true, Cc: ['naresh@example.test', 'kiran@example.test'] });
    // The existing entry keeps its display name; the new one is appended.
    expect(now().CcRecipients).toEqual([addr('naresh@example.test'), { EmailAddress: { Address: 'kiran@example.test' } }]);
    await h.close();
  });

  it('does not duplicate an address already on the Cc line (any case)', async () => {
    const { client } = withCc();
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', addCc: ['NARESH@example.test', 'kiran@example.test'] });
    expect(res.Cc).toEqual(['naresh@example.test', 'kiran@example.test']);
    await h.close();
  });

  it('removes one cc and leaves the others', async () => {
    const { client } = stub({ messages: [draft({ CcRecipients: [addr('naresh@example.test'), addr('kiran@example.test')] })] });
    const h = await harness(client);
    const res = await confirmed(h, { id: 'd1', removeCc: ['Kiran@example.test'] });
    expect(res).toMatchObject({ updated: true, Cc: ['naresh@example.test'] });
    await h.close();
  });

  it('refuses to remove an address that is not on the Cc line, naming the current Cc', async () => {
    const { client, write } = withCc();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_draft', { id: 'd1', removeCc: ['nobody@example.test'] });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/nobody@example\.test/);
    expect(text(res)).toMatch(/naresh@example\.test/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('refuses cc together with addCc or removeCc', async () => {
    const { client, write } = withCc();
    const h = await harness(client);
    const res = await h.callTool('outlook_update_draft', { id: 'd1', cc: ['a@example.test'], addCc: ['b@example.test'] });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/addCc|removeCc/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });
});

describe('outlook_send_draft', () => {
  const send = (h: TestHarness, args: Obj) => h.callTool('outlook_send_draft', args);
  const sendJson = async (h: TestHarness, args: Obj) => parseToolResult<Result>(await send(h, args));
  const sendConfirmed = async (h: TestHarness, args: Obj) => {
    const first = await sendJson(h, args);
    expect(first.status).toBe('confirmation-required');
    return sendJson(h, { ...args, confirmToken: first.confirmToken });
  };

  it('previews exactly what will go out, to whom, and sends nothing without a token', async () => {
    const { client, write } = stub({ messages: [draft({ CcRecipients: [addr('cc@example.test')] })] });
    const h = await harness(client);
    const res = await sendJson(h, { id: 'd1' });
    expect(res.status).toBe('confirmation-required');
    const shown = JSON.stringify(res);
    expect(shown).toMatch(/Quarterly numbers/);
    expect(shown).toMatch(/ann@example\.test/);
    expect(shown).toMatch(/cc@example\.test/);
    expect(shown).toMatch(/Old text the user typed in Outlook/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('sends after confirmation and verifies the draft has left Drafts', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await sendConfirmed(h, { id: 'd1' });
    expect(write.mock.calls).toEqual([['POST', '/me/messages/d1/send']]);
    expect(res).toMatchObject({ sent: true, subject: 'Quarterly numbers', recipients: ['ann@example.test'] });
    expect(res.warning).toBeUndefined();
    await h.close();
  });

  it('waits for Outlook to finish sending, as it does live a moment after the 202', async () => {
    const { client, get, write } = stub({ persist: false });
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    let reads = 0;
    get.mockImplementation(async (p: string) => {
      if (write.mock.calls.length > 0 && ++reads >= 3) throw new ApiError(404, 'Outlook 404: ErrorItemNotFound');
      return real(p);
    });
    const h = await harness(client);
    const res = await sendConfirmed(h, { id: 'd1' });
    expect(res).toMatchObject({ sent: true });
    expect(res.warning).toBeUndefined();
    await h.close();
  });

  it('never says "not sent" after Outlook accepted it: a lingering draft means check Sent Items, not resend', async () => {
    const { client } = stub({ persist: false });
    const h = await harness(client);
    const res = await sendConfirmed(h, { id: 'd1' });
    expect(res).toMatchObject({ sent: true, verified: false });
    expect(res.warning).toMatch(/Sent Items/);
    expect(res.warning).toMatch(/not send it again|do not resend|before sending again/i);
    await h.close();
  });

  it('refuses a draft with no recipients, pointing at outlook_update_draft', async () => {
    const { client, write } = stub({ messages: [draft({ ToRecipients: [] })] });
    const h = await harness(client);
    const res = await send(h, { id: 'd1' });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/no recipients/i);
    expect(text(res)).toMatch(/outlook_update_draft/);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('refuses a message that is not a draft', async () => {
    const { client, write } = stub();
    const h = await harness(client);
    const res = await send(h, { id: 'sent1' });
    expect(res.isError).toBe(true);
    expect(text(res)).toMatch(/not a draft/i);
    expect(write).not.toHaveBeenCalled();
    await h.close();
  });

  it('reports an unknown outcome and says to check Sent Items before resending', async () => {
    const { client } = stub({ writeError: new WriteOutcomeUnknownError('Outlook', 'POST', { timeoutMs: 60_000 }) });
    const h = await harness(client);
    const res = await sendConfirmed(h, { id: 'd1' });
    expect(res.status).toBe('unknown');
    expect(res.warning).toMatch(/Sent Items/);
    await h.close();
  });

  it('says it could not check when the send lands but Drafts cannot be re-read', async () => {
    const { client, get, write } = stub();
    const real = get.getMockImplementation() as (p: string) => Promise<unknown>;
    get.mockImplementation(async (p: string) => {
      if (write.mock.calls.length > 0) throw new ApiError(503, 'Outlook 503');
      return real(p);
    });
    const h = await harness(client);
    const res = await sendConfirmed(h, { id: 'd1' });
    expect(res).toMatchObject({ sent: true });
    expect(res.warning).toMatch(/before retrying/i);
    await h.close();
  });
});
