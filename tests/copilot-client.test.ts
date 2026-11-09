import { describe, expect, it, afterEach } from 'vitest';
import { CopilotClient } from '../src/copilot/client.js';
import { startStub, delta, completion, fakeToken, type StubServer } from './_copilot-stub.js';

let stub: StubServer | undefined;
afterEach(async () => {
  await stub?.close();
  stub = undefined;
});

/** A CopilotClient wired to the stub endpoint with an injected token. */
function client(origin: string) {
  return new CopilotClient({
    env: { OUTLOOK_COPILOT_ENDPOINT: origin } as NodeJS.ProcessEnv,
    tokenClient: { getToken: async () => fakeToken('OID', 'TID'), invalidate: () => {} },
  });
}

describe('CopilotClient', () => {
  it('asks over a real session and returns the answer', async () => {
    stub = await startStub({ respond: () => [delta('hi '), delta('there'), completion()] });
    const result = await client(stub.origin).ask('question');
    expect(result.text).toBe('hi there');
    expect(stub.last().url).toContain('/m365Copilot/Chathub/OID@TID');
  });

  it('threads turns on one conversation until reset() starts a new one', async () => {
    stub = await startStub();
    const c = client(stub.origin);
    const first = c.conversationId;
    await c.ask('one');
    expect(c.conversationId).toBe(first); // same conversation across turns
    c.reset();
    expect(c.conversationId).not.toBe(first);
  });
});
