import { describe, expect, it, afterEach, vi } from 'vitest';
import { CopilotSession, DisengagedError, decodeJwt, cleanAnswer } from '../src/copilot/session.js';
import {
  startStub,
  delta,
  botMessage,
  disengaged,
  throttling,
  completion,
  fakeToken,
  type StubServer,
} from './_copilot-stub.js';

let stub: StubServer | undefined;
afterEach(async () => {
  await stub?.close();
  stub = undefined;
});

function session(endpoint: string, token = fakeToken()) {
  return new CopilotSession({ getToken: async () => token, endpoint });
}

describe('CopilotSession', () => {
  it('handshakes, sends the chat AND metrics frames together, and streams the answer', async () => {
    stub = await startStub();
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('Hello world');

    const frames = stub.last().frames;
    expect(frames[0]).toEqual({ protocol: 'json', version: 1 });
    const chat = frames.find((f) => f.type === 4 && f.target === 'chat')!;
    const metrics = frames.find((f) => f.target === 'Metrics')!;
    // The Metrics frame is mandatory and must arrive with the chat frame.
    expect(chat).toBeTruthy();
    expect(metrics).toBeTruthy();
    const arg = (chat.arguments as Record<string, unknown>[])[0]!;
    expect(arg.tone).toBe('Magic');
    expect(arg.isStartOfSession).toBe(true);
    // WorkSearch is what grounds the turn on the mailbox.
    expect(JSON.stringify(arg.gpts)).toContain('WorkSearch');
  });

  it('puts the token in the URL and addresses the hub by oid@tid', async () => {
    stub = await startStub();
    await session(stub.origin, fakeToken('OID', 'TID')).run('hi');
    expect(stub.last().url).toContain('/m365Copilot/Chathub/OID@TID');
    expect(stub.last().url).toContain('access_token=');
    expect(stub.last().url).toContain('scenario=OutlookSidepaneWorkCopilot');
  });

  it('prefers the longer of the streamed deltas and the snapshot message', async () => {
    stub = await startStub({
      respond: () => [delta('Hel'), botMessage('Hello, full snapshot'), completion()],
    });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('Hello, full snapshot');
  });

  it('reports the per-conversation quota when the server sends it', async () => {
    stub = await startStub({ respond: () => [delta('ok'), throttling(3, 600), completion()] });
    const result = await session(stub.origin).run('hi');
    expect(result.throttle).toEqual({ current: 3, max: 600 });
  });

  it('ignores bot messages that carry a messageType (control traffic, not the answer)', async () => {
    stub = await startStub({
      respond: () => [
        botMessage('Searching…', { messageType: 'Progress' }),
        delta('real answer'),
        completion(),
      ],
    });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('real answer');
  });

  it('raises DisengagedError on a Disengaged refusal, not an empty success', async () => {
    stub = await startStub({ respond: () => [disengaged(), completion()] });
    await expect(session(stub.origin).run('hi')).rejects.toBeInstanceOf(DisengagedError);
  });

  it('answers a keepalive ping and still completes', async () => {
    stub = await startStub({ pingFirst: true, respond: () => [delta('pong-ok'), completion()] });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('pong-ok');
    // The pong is in flight to the stub when run() resolves; wait for it to land.
    await vi.waitFor(() => expect(stub!.last().frames.some((f) => f.type === 6)).toBe(true));
  });

  it('rejects when the socket closes before a completion frame', async () => {
    stub = await startStub({ respond: () => [delta('partial')] }); // no completion
    const run = session(stub.origin).run('hi');
    // Close the connection from the server side without ever completing the turn.
    setTimeout(() => stub?.last()?.socket.close(), 50);
    await expect(run).rejects.toThrow(/closed before the turn completed/i);
  });

  it('rejects a completion frame that carries an error', async () => {
    stub = await startStub({ respond: () => [completion('BadRequest')] });
    await expect(session(stub.origin).run('hi')).rejects.toThrow(/rejected the turn/i);
  });

  it('threads turns by conversation id, and reset() starts a fresh one', async () => {
    stub = await startStub();
    const s = session(stub.origin);
    const first = s.conversationId;
    await s.run('one');
    await s.run('two');
    // Second turn is not the start of a session and reuses the conversation.
    const secondChat = stub.connections[1]!.frames.find((f) => f.target === 'chat')!;
    expect((secondChat.arguments as Record<string, unknown>[])[0]!.isStartOfSession).toBe(false);
    s.reset();
    expect(s.conversationId).not.toBe(first);
  });

  it('strips citation control tokens from the answer, keeping the prose', () => {
    const raw = '\uE200You have \uE201170 unread. citeturn1search23turn1search7 Done.\u30102-9c5ade\u3011';
    const clean = cleanAnswer(raw);
    expect(clean).toBe('You have 170 unread. Done.');
    expect(clean).not.toMatch(/cite|turn1search|[\uE000-\uF8FF]|\u3010/u);
  });

  it('returns already-clean prose untouched', () => {
    expect(cleanAnswer('Just three bullet points, nothing fancy.')).toBe('Just three bullet points, nothing fancy.');
  });

  it('preserves leading indentation of a quoted code block while collapsing mid-line gaps', () => {
    const raw = 'Here is the snippet:\n```\nfunction f() {\n    if (x) {\n\t\treturn 1;\n    }\n}\n```\nThat  was  it.';
    expect(cleanAnswer(raw)).toBe(
      'Here is the snippet:\n```\nfunction f() {\n    if (x) {\n\t\treturn 1;\n    }\n}\n```\nThat was it.',
    );
  });

  it('collapses the double space a removed citation token leaves mid-line', () => {
    expect(cleanAnswer('Two items citeturn1search3  remain.')).toBe('Two items remain.');
  });

  it('cleans citation tokens that arrive in the streamed answer', async () => {
    stub = await startStub({
      respond: () => [delta('Reviewed. citeturn1search3 '), delta('All good.\u30101-abc123\u3011'), completion()],
    });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('Reviewed. All good.');
  });

  it('decodeJwt rejects malformed tokens', () => {
    expect(() => decodeJwt('abc')).toThrow(/expected a JWT/i);
    const notJson = `eyJhbGciOiJub25lIn0.${Buffer.from('not-json').toString('base64url')}.sig`;
    expect(() => decodeJwt(notJson)).toThrow(/base64url JSON/i);
    const noClaims = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ sub: 'x' })).toString('base64url')}.sig`;
    expect(() => decodeJwt(noClaims)).toThrow(/oid\/tid/i);
  });

  it('completes on a clean type-7 frame', async () => {
    stub = await startStub({ respond: () => [delta('seven'), { type: 7 }] });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('seven');
  });

  it('rejects a type-7 frame that carries an error', async () => {
    stub = await startStub({ respond: () => [{ type: 7, error: 'closed' }] });
    await expect(session(stub.origin).run('hi')).rejects.toThrow(/closed the connection/i);
  });

  it('reads throttle and contentOrigin off a type-2 stream item', async () => {
    stub = await startStub({
      respond: () => [
        delta('answer'),
        { type: 2, item: { throttling: { numUserMessagesInConversation: 5, maxNumUserMessagesInConversation: 600 }, messages: [{ author: 'bot', contentOrigin: 'DeepLeo' }] } },
        completion(),
      ],
    });
    const result = await session(stub.origin).run('hi');
    expect(result.text).toBe('answer');
    expect(result.throttle).toEqual({ current: 5, max: 600 });
    expect(result.contentOrigin).toBe('DeepLeo');
  });

  it('aborts a turn when its signal fires', async () => {
    stub = await startStub({ respond: () => [] }); // never completes
    const ac = new AbortController();
    const run = session(stub.origin).run('hi', { signal: ac.signal });
    setTimeout(() => ac.abort(), 30);
    await expect(run).rejects.toThrow(/aborted/i);
  });
});
