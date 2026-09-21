import { describe, expect, it, vi } from 'vitest';
import { OutlookClient, DEFAULT_API_BASE } from '../src/client.js';
import { stripBearer, raceCaptures } from '../src/auth-fetchproxy.js';

/** A JWT whose `exp` is `secondsFromNow` out. Signature is irrelevant here. */
function jwt(secondsFromNow: number): string {
  const b64 = (o: unknown) =>
    Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({ exp: Math.floor(Date.now() / 1000) + secondsFromNow })}.x`;
}

/**
 * Test env.
 *
 * Every client here gets an EXPLICIT env object, which means `tests/_setup.ts`
 * cannot protect it: the client reads the object it was handed, not
 * `process.env`. Without `OUTLOOK_TOKEN_CACHE: 'false'` the cache defaults to
 * ON and resolves to the developer's real `~/.office-outlook-mcp`, so one test
 * writes a token there and the next silently loads it instead of capturing.
 */
function env(overrides: Record<string, string> = {}): NodeJS.ProcessEnv {
  return { OUTLOOK_TOKEN_CACHE: 'false', ...overrides };
}

/**
 * The `Prefer` header of the nth request.
 *
 * Goes through `init.headers` deliberately. `new Headers(init)` does NOT throw
 * — it treats each own property of the RequestInit as a header, yielding
 * entries named `method`/`headers`/`signal` — so `.get('prefer')` returns null
 * for every call and a `toBeNull()` assertion passes for entirely the wrong
 * reason.
 */
function preferOf(mock: { mock: { calls: unknown[] } }, n: number): string | null {
  const call = mock.mock.calls[n] as [unknown, RequestInit] | undefined;
  if (!call) throw new Error(`no request at index ${n}`);
  return new Headers(call[1]?.headers).get('prefer');
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

describe('credential configuration', () => {
  it('does not throw when nothing is configured, so the server can still boot', () => {
    expect(
      () => new OutlookClient({ env: env({ OUTLOOK_DISABLE_FETCHPROXY: '1' }) }),
    ).not.toThrow();
  });

  it('defers the config error to the first call, with an actionable hint', async () => {
    const c = new OutlookClient({ env: env({ OUTLOOK_DISABLE_FETCHPROXY: '1' }) });
    await expect(c.get('/me')).rejects.toThrow(/No Outlook credential is configured/);
    await expect(c.get('/me')).rejects.toMatchObject({
      hint: expect.stringContaining('OUTLOOK_ACCESS_TOKEN'),
    });
  });

  it('reports a non-secret token source', () => {
    expect(new OutlookClient({ env: env({ OUTLOOK_ACCESS_TOKEN: 'x' }) }).tokenSource).toBe(
      'OUTLOOK_ACCESS_TOKEN',
    );
    expect(new OutlookClient({ env: env() }).tokenSource).toBe('browser capture');
  });

  it('uses the documented API base by default and honours an override', () => {
    expect(new OutlookClient({ env: env() }).apiBase).toBe(DEFAULT_API_BASE);
    expect(
      new OutlookClient({ env: env({ OUTLOOK_API_BASE: 'https://example.test/v9' }) }).apiBase,
    ).toBe('https://example.test/v9');
  });
});

describe('requests', () => {
  it('sends the captured token as a bearer and hits the given path', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ EmailAddress: 'a@b.c' }));
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(3600),
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    await expect(c.get('/me')).resolves.toMatchObject({ EmailAddress: 'a@b.c' });

    const [url, init] = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(String(url)).toBe(`${DEFAULT_API_BASE}/me`);
    const headers = new Headers(init.headers);
    expect(headers.get('authorization')).toMatch(/^Bearer /);
  });

  it('asks for a plain-text body only when requested', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({}));
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(3600),
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });

    await c.get('/me/messages/1');
    expect(preferOf(fetchImpl, 0)).toBeNull();

    await c.get('/me/messages/1', { text: true });
    expect(preferOf(fetchImpl, 1)).toBe('outlook.body-content-type="text"');
  });

  it('combines several Prefer values into one header', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({}));
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(3600),
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    await c.get('/me/events/1', { text: true, prefer: 'outlook.timezone="UTC"' });
    expect(preferOf(fetchImpl, 0)).toBe(
      'outlook.body-content-type="text", outlook.timezone="UTC"',
    );
  });

  it('refuses to follow a pagination link off the API origin', async () => {
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(3600),
      fetchImpl: (async () => jsonResponse({})) as unknown as typeof fetch,
    });
    await expect(c.getAbsolute('https://evil.example.com/steal')).rejects.toThrow(
      /Refusing to follow a link/,
    );
  });

  it('follows a same-origin pagination link', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ value: [] }));
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(3600),
      fetchImpl: fetchImpl as unknown as typeof fetch,
    });
    await c.getAbsolute(`${DEFAULT_API_BASE}/me/messages?$skip=10`);
    expect(String((fetchImpl.mock.calls[0] as never[])[0])).toContain('$skip=10');
  });
});

describe('token lifecycle', () => {
  it('captures once and reuses the token across calls', async () => {
    const capture = vi.fn(async () => jwt(3600));
    const c = new OutlookClient({
      env: env(),
      captureToken: capture,
      fetchImpl: (async () => jsonResponse({})) as unknown as typeof fetch,
    });
    await c.get('/me');
    await c.get('/me');
    expect(capture).toHaveBeenCalledTimes(1);
  });

  it('derives expiry from the token rather than guessing', async () => {
    const c = new OutlookClient({
      env: env(),
      captureToken: async () => jwt(7200),
      fetchImpl: (async () => jsonResponse({})) as unknown as typeof fetch,
    });
    await c.get('/me');
    const left = (c.tokenExpiresAt() ?? 0) - Date.now();
    expect(left).toBeGreaterThan(7000_000);
    expect(left).toBeLessThan(7300_000);
  });

  it('falls back to a short TTL for an opaque, non-JWT token', async () => {
    const c = new OutlookClient({
      env: env({ OUTLOOK_ACCESS_TOKEN: 'opaque-not-a-jwt' }),
      fetchImpl: (async () => jsonResponse({})) as unknown as typeof fetch,
    });
    await c.get('/me');
    const left = (c.tokenExpiresAt() ?? 0) - Date.now();
    expect(left).toBeGreaterThan(0);
    expect(left).toBeLessThanOrEqual(10 * 60 * 1000);
  });

  it('never reaches for the browser to refresh a directly-supplied token', async () => {
    // Regression: a 401 used to send OUTLOOK_ACCESS_TOKEN users to the bridge,
    // which silently swaps the operator's chosen credential for whatever
    // account a tab is signed into — and hangs for the capture window on a
    // machine with no bridge at all.
    const capture = vi.fn(async () => jwt(3600));
    const c = new OutlookClient({
      env: env({ OUTLOOK_ACCESS_TOKEN: 'opaque-not-a-jwt' }),
      captureToken: capture,
      fetchImpl: (async () => jsonResponse({ error: 'nope' }, 401)) as unknown as typeof fetch,
    });
    await expect(c.get('/me')).rejects.toThrow();
    expect(capture).not.toHaveBeenCalled();
  });
});

describe('stripBearer', () => {
  it('removes a Bearer prefix in any casing and trims', () => {
    expect(stripBearer('Bearer abc')).toBe('abc');
    expect(stripBearer('bearer  abc ')).toBe('abc');
    expect(stripBearer('  abc  ')).toBe('abc');
  });
});

describe('bridge trust boundary', () => {
  it('trusts exactly the hosts it captures from, and no wider', async () => {
    // Was the apex pair `cloud.microsoft` / `office.com`, which grants the
    // extension every subdomain of both — the whole of Microsoft 365 — to read
    // one header off two known hosts. It is also what sent the pairing flow to
    // `m365.cloud.microsoft/chat`, a tab nothing here needs. Deriving the trust
    // set from the capture declarations keeps the two from drifting apart.
    const { CAPTURE_HOSTS, TRUST_DOMAINS } = await import('../src/auth-fetchproxy.js');
    expect([...TRUST_DOMAINS].sort()).toEqual([...CAPTURE_HOSTS].sort());
    for (const d of TRUST_DOMAINS) {
      expect(d).toMatch(/^outlook\./);
    }
  });
});

describe('raceCaptures', () => {
  it('returns the first host that answers and ignores the other failing', async () => {
    const slowFail = new Promise<string>((_, rej) => setTimeout(() => rej(new Error('nope')), 50));
    await expect(raceCaptures([Promise.resolve('tok'), slowFail], 1_000)).resolves.toBe('tok');
  });

  it('gives up ONE window after it started, not one window per declared host', async () => {
    // Measured live 2026-09-20: the two declared hosts are documented as
    // "raced", but the bridge serializes them, so the wait was ~2x the
    // configured window — 5s produced 14.4s and 20s produced 36s. At the
    // documented 30s default that is past 60s, which is the default request
    // timeout in an MCP client: the capture blew the host's deadline before it
    // could report its own error. The knob must bound the TOTAL wait.
    vi.useFakeTimers();
    try {
      const never = () => new Promise<string>(() => {});
      const p = raceCaptures([never(), never()], 20_000);
      const assertion = expect(p).rejects.toThrow(/20s/);
      await vi.advanceTimersByTimeAsync(20_100);
      await assertion;
    } finally {
      vi.useRealTimers();
    }
  });

  it('reports what each host said when all of them fail', async () => {
    const p = raceCaptures(
      [Promise.reject(new Error('cloud said no')), Promise.reject(new Error('office said no'))],
      1_000,
    );
    await expect(p).rejects.toThrow(/cloud said no.*office said no/s);
  });
});
