import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerCopilotTools } from '../src/tools/copilot.js';
import { DisengagedError, type TurnResult } from '../src/copilot/session.js';
import type { CopilotClient } from '../src/copilot/client.js';
import type { OutlookClient } from '../src/client.js';

/** A stub CopilotClient: the registrar only needs ask/reset/conversationId. */
function stub(
  ask: (prompt: string, opts: { signal?: AbortSignal }) => Promise<TurnResult>,
  conversationId = 'conv-1',
) {
  const reset = vi.fn();
  const client = { ask: vi.fn(ask), reset, conversationId } as unknown as CopilotClient;
  return { client, reset, ask: client.ask as ReturnType<typeof vi.fn> };
}

/** A stub OutlookClient whose MailboxSettings carry a time zone (or none / an error). */
function outlookStub(tz: string | undefined | 'throw' = 'Eastern Standard Time'): OutlookClient {
  return {
    get: vi.fn(async (p: string) => {
      if (tz === 'throw') throw new Error('no mail token');
      return p === '/me/MailboxSettings' ? { TimeZone: tz } : {};
    }),
  } as unknown as OutlookClient;
}

const harness = (client: CopilotClient, outlook: OutlookClient = outlookStub()) =>
  createTestHarness((s: McpServer) => registerCopilotTools(s, client, outlook));

type Envelope = { untrusted_content?: boolean; answer?: string; conversationId?: string; quota?: { used: number; max: number }; contentOrigin?: string };

describe('outlook_ask_copilot', () => {
  it('returns Copilot\'s answer inside the untrusted-content envelope, with quota', async () => {
    const { client, ask } = stub(async () => ({
      text: 'You have 3 urgent emails.',
      throttle: { current: 2, max: 600 },
      contentOrigin: 'DeepLeo',
    }));
    const h = await harness(client);
    const res = parseToolResult<Envelope>(await h.callTool('outlook_ask_copilot', { prompt: 'what is urgent?' }));
    expect(res.untrusted_content).toBe(true);
    expect(res).toMatchObject({
      answer: 'You have 3 urgent emails.',
      conversationId: 'conv-1',
      quota: { used: 2, max: 600 },
      contentOrigin: 'DeepLeo',
    });
    // The user's question is carried verbatim, wrapped in a time-zone context preamble.
    const sent = ask.mock.calls[0]?.[0] as string;
    expect(sent).toContain('what is urgent?');
    expect(sent).toMatch(/Eastern Standard Time/);
    expect(ask).toHaveBeenCalledWith(expect.any(String), expect.objectContaining({ signal: expect.anything() }));
    await h.close();
  });

  it('prepends the mailbox time zone so Copilot labels times, and degrades when it is unknown', async () => {
    const { client, ask } = stub(async () => ({ text: 'ok' }));
    const h = await harness(client, outlookStub('Pacific Standard Time'));
    await h.callTool('outlook_ask_copilot', { prompt: 'when did X email me?' });
    expect(ask.mock.calls[0]?.[0]).toMatch(/Pacific Standard Time/);

    const noTz = stub(async () => ({ text: 'ok' }));
    const h2 = await harness(noTz.client, outlookStub('throw'));
    await h2.callTool('outlook_ask_copilot', { prompt: 'when?' });
    // No mailbox token: still a time-zone instruction, just without the named zone.
    expect(noTz.ask.mock.calls[0]?.[0]).toMatch(/time zone/i);
    expect(noTz.ask.mock.calls[0]?.[0]).toContain('when?');
    await h.close();
    await h2.close();
  });

  it('adds a verbatim directive only when asked', async () => {
    const { client, ask } = stub(async () => ({ text: 'ok' }));
    const h = await harness(client);
    await h.callTool('outlook_ask_copilot', { prompt: 'quote my note' });
    expect(ask.mock.calls[0]?.[0]).not.toMatch(/verbatim/i);
    await h.callTool('outlook_ask_copilot', { prompt: 'quote my note', verbatim: true });
    expect(ask.mock.calls[1]?.[0]).toMatch(/verbatim/i);
    await h.close();
  });

  it('continues the conversation by default and only resets when asked', async () => {
    const { client, reset } = stub(async () => ({ text: 'ok' }));
    const h = await harness(client);
    await h.callTool('outlook_ask_copilot', { prompt: 'a' });
    expect(reset).not.toHaveBeenCalled();
    await h.callTool('outlook_ask_copilot', { prompt: 'b', newConversation: true });
    expect(reset).toHaveBeenCalledTimes(1);
    await h.close();
  });

  it('reports a Disengaged refusal as a non-error result that says not to retry', async () => {
    const { client } = stub(async () => {
      throw new DisengagedError();
    });
    const h = await harness(client);
    const raw = await h.callTool('outlook_ask_copilot', { prompt: 'something it dislikes' });
    expect(raw.isError).toBeFalsy();
    const res = parseToolResult<{ disengaged?: boolean; note?: string }>(raw);
    expect(res.disengaged).toBe(true);
    expect(res.note).toMatch(/not a rate limit/i);
    await h.close();
  });

  it('surfaces a missing-credential error as a tool error', async () => {
    const { client } = stub(async () => {
      throw new McpToolError('No signed-in Microsoft 365 Copilot account.', {
        hint: 'Sign in once with `opencode-m365 login`.',
        kind: 'no_credential',
      });
    });
    const h = await harness(client);
    const raw = await h.callTool('outlook_ask_copilot', { prompt: 'hi' });
    expect(raw.isError).toBe(true);
    expect(JSON.stringify(raw.content)).toMatch(/opencode-m365 login/);
    await h.close();
  });
});
