import { describe, expect, it, vi } from 'vitest';
import { createCopilotTokenClient, copilotCacheFile, type MsalLike } from '../src/copilot/auth.js';

/** A fake MSAL client: one account, a token with a far-future expiry. */
function fakeApp(over: Partial<{ accounts: unknown[]; token: unknown; throws: boolean }> = {}): {
  app: MsalLike;
  acquire: ReturnType<typeof vi.fn>;
} {
  const accounts = over.accounts ?? [{ username: 'me@example.test' }];
  const token =
    over.token === undefined ? { accessToken: 'tok-1', expiresOn: new Date(Date.now() + 3_600_000) } : over.token;
  const acquire = vi.fn(async () => {
    if (over.throws) throw new Error('interaction_required');
    return token as { accessToken?: string; expiresOn?: Date | null };
  });
  const app: MsalLike = {
    getTokenCache: () => ({ getAllAccounts: async () => accounts }),
    acquireTokenSilent: acquire,
  };
  return { app, acquire };
}

describe('createCopilotTokenClient', () => {
  it('returns the Sydney access token and caches it (no second acquire)', async () => {
    const { app, acquire } = fakeApp();
    const client = createCopilotTokenClient({ app });
    expect(await client.getToken()).toBe('tok-1');
    expect(await client.getToken()).toBe('tok-1');
    expect(acquire).toHaveBeenCalledTimes(1);
  });

  it('re-acquires after invalidate()', async () => {
    const { app, acquire } = fakeApp();
    const client = createCopilotTokenClient({ app });
    await client.getToken();
    client.invalidate();
    await client.getToken();
    expect(acquire).toHaveBeenCalledTimes(2);
  });

  it('points the user at `opencode-m365 login` when there is no account', async () => {
    const { app } = fakeApp({ accounts: [] });
    const client = createCopilotTokenClient({ app });
    await expect(client.getToken()).rejects.toMatchObject({ kind: 'no_credential', hint: expect.stringMatching(/opencode-m365 login/) });
  });

  it('surfaces a silent-refresh failure as a credential error, not a raw throw', async () => {
    const { app } = fakeApp({ throws: true });
    const client = createCopilotTokenClient({ app });
    await expect(client.getToken()).rejects.toMatchObject({ kind: 'no_credential', hint: expect.stringMatching(/opencode-m365 login/) });
  });

  it('treats a result with no access token as a credential error', async () => {
    const { app } = fakeApp({ token: { accessToken: undefined } });
    const client = createCopilotTokenClient({ app });
    await expect(client.getToken()).rejects.toMatchObject({ kind: 'no_credential', hint: expect.stringMatching(/opencode-m365 login/) });
  });

  it('with no injected client, loads MSAL over the cache file and reports no account', async () => {
    // No `app` injected: this drives the real lazy import + cache-plugin read
    // (a missing file is swallowed) and the empty-account path, all offline.
    const client = createCopilotTokenClient({ cacheFile: '/nonexistent/office-outlook-mcp-copilot/msal.json' });
    await expect(client.getToken()).rejects.toMatchObject({ kind: 'no_credential' });
  });

  it('defaults the cache to the opencode-m365 location, overridable by env', async () => {
    expect(copilotCacheFile({ HOME: '/home/x' } as NodeJS.ProcessEnv)).toContain('opencode-copilot');
    expect(copilotCacheFile({ OUTLOOK_COPILOT_MSAL_CACHE: '/tmp/c.json' } as NodeJS.ProcessEnv)).toBe('/tmp/c.json');
  });
});
