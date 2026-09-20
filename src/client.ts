/**
 * The Outlook REST v2.0 client.
 *
 * Auth is a bearer token whose only source is the user's browser (see
 * `auth-fetchproxy.ts`), or an `OUTLOOK_ACCESS_TOKEN` supplied directly. Once
 * held, every request is a plain Node `fetch` — the bridge is never on the hot
 * path.
 */
import {
  createApiClient,
  decodeJwtExp,
  McpToolError,
  parseBoolEnv,
  readEnvVar,
  type ApiClient,
} from '@chrischall/mcp-utils';
import { TokenManager, type BearerTokens } from '@chrischall/mcp-utils/session';
import { createTokenCache, reportCacheWriteFailure } from './token-cache.js';

/**
 * Load the bridge bootstrap lazily.
 *
 * `auth-fetchproxy.ts` pulls in `@chrischall/mcp-utils/fetchproxy`, which
 * imports the OPTIONAL `@fetchproxy/server` peer. That package is
 * esbuild-`--external`, and the `.mcpb` bundle ships no `node_modules` — so a
 * top-level import here would throw `ERR_MODULE_NOT_FOUND` the moment a host
 * spawns the bundled server, before it can answer `initialize`. The host then
 * reports only "Server transport closed unexpectedly".
 *
 * Deferring it means the default path never touches the package, and a user who
 * supplies `OUTLOOK_ACCESS_TOKEN` never needs it installed at all.
 */
async function captureTokenLazily(): Promise<string> {
  const mod = await import('./auth-fetchproxy.js');
  return mod.captureTokenViaFetchproxy();
}

/**
 * The API the captured token is minted for.
 *
 * Note this is `outlook.office.com` even when the signed-in tab is on
 * `outlook.cloud.microsoft` — the token's `aud` claim names this host, and it
 * is the host that answers. Verified live 2026-09-20.
 */
export const DEFAULT_API_BASE = 'https://outlook.office.com/api/v2.0';

/**
 * Fallback lifetime for a token whose `exp` we cannot read.
 *
 * Real Outlook tokens are JWTs carrying `exp` (~25h out), so this is only
 * reached for a hand-supplied opaque `OUTLOOK_ACCESS_TOKEN`. Deliberately
 * short: being wrong in this direction costs one extra 401-and-replay, while
 * being wrong the other way makes every call fail until a restart.
 */
const UNKNOWN_TOKEN_TTL_MS = 10 * 60 * 1000;

/**
 * Sentinel standing in for a refresh token this flow does not have.
 *
 * Outlook Web's token comes from MSAL's hidden-iframe flow; there is no refresh
 * grant we can drive. "Refreshing" therefore means capturing a new one. The
 * sentinel exists because `TokenManager` only calls `refresh` when a refresh
 * token is present, and without it an expired token would have no path back
 * short of restarting the process.
 */
const RECAPTURE_SENTINEL = 'fetchproxy:recapture';

function toBearerTokens(accessToken: string): BearerTokens {
  let expiresAt: number;
  try {
    // decodeJwtExp returns seconds; a non-JWT throws.
    expiresAt = decodeJwtExp(accessToken) * 1000;
  } catch {
    expiresAt = Date.now() + UNKNOWN_TOKEN_TTL_MS;
  }
  return { accessToken, refreshToken: RECAPTURE_SENTINEL, expiresAt };
}

export interface OutlookClientOptions {
  env?: NodeJS.ProcessEnv;
  /** Injected for tests so no suite ever touches the real bridge. */
  captureToken?: () => Promise<string>;
  fetchImpl?: typeof fetch;
}

export class OutlookClient {
  readonly #baseUrl: string;
  /**
   * Deferred configuration error.
   *
   * The constructor never throws, so the server still boots (and answers a
   * host's install-time `tools/list` probe) with no credential present; the
   * error surfaces on the first tool call instead.
   */
  readonly #configError: McpToolError | null = null;
  #api: ApiClient | null = null;
  readonly #env: NodeJS.ProcessEnv;
  readonly #capture: () => Promise<string>;
  readonly #fetchImpl: typeof fetch;
  #tokens: TokenManager | null = null;
  #tokenSource = 'unconfigured';

  constructor(opts: OutlookClientOptions = {}) {
    this.#env = opts.env ?? process.env;
    this.#capture = opts.captureToken ?? captureTokenLazily;
    // A receiver-safe wrapper, never the bare global. Storing `fetch` itself
    // and calling it as `this.#fetchImpl(...)` binds `this` to the client, and
    // older undici (Node 18-20, which several MCP hosts still bundle) throws
    // `Illegal invocation` for any receiver that is not `globalThis`.
    this.#fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));
    this.#baseUrl = readEnvVar('OUTLOOK_API_BASE', { env: this.#env }) ?? DEFAULT_API_BASE;

    const direct = readEnvVar('OUTLOOK_ACCESS_TOKEN', { env: this.#env });
    const bridgeOff = parseBoolEnv('OUTLOOK_DISABLE_FETCHPROXY', { env: this.#env });
    if (direct === undefined && bridgeOff) {
      this.#configError = new McpToolError(
        'No Outlook credential is configured.',
        {
          hint:
            'Set OUTLOOK_ACCESS_TOKEN, or unset OUTLOOK_DISABLE_FETCHPROXY so the ' +
            'token can be captured from a signed-in Outlook browser tab.',
        },
      );
    }
    this.#tokenSource = direct !== undefined ? 'OUTLOOK_ACCESS_TOKEN' : 'browser capture';
  }

  /** Non-secret description of where the credential comes from. */
  get tokenSource(): string {
    return this.#tokenSource;
  }

  #requireConfigured(): void {
    if (this.#configError) throw this.#configError;
  }

  #tokenManager(): TokenManager {
    if (this.#tokens) return this.#tokens;

    const direct = readEnvVar('OUTLOOK_ACCESS_TOKEN', { env: this.#env });
    const persistence = direct === undefined ? createTokenCache(this.#env) : null;
    const mint = async (): Promise<BearerTokens> =>
      toBearerTokens(direct ?? (await this.#capture()));

    this.#tokens = new TokenManager({
      initial: mint,
      // There is no refresh grant — re-capture instead. The sentinel argument
      // is ignored by design.
      //
      // But a token supplied as OUTLOOK_ACCESS_TOKEN must NOT be "refreshed" by
      // reaching for the browser: that silently swaps the credential the
      // operator chose for whichever account a tab happens to be signed into,
      // and on a machine with no bridge it just hangs for the capture window on
      // every 401. Say what is wrong instead.
      refresh: async () => {
        if (direct !== undefined) {
          throw new McpToolError('OUTLOOK_ACCESS_TOKEN was rejected by Outlook.', {
            hint:
              'A directly-supplied token cannot be refreshed automatically. Replace ' +
              'OUTLOOK_ACCESS_TOKEN with a current one, or unset it to capture from a ' +
              'signed-in Outlook browser tab instead.',
          });
        }
        return toBearerTokens(await this.#capture());
      },
      // A failed cache write must not fail the call that triggered it: the
      // token is in hand and re-capturable, so a read-only data dir should cost
      // the next cold start a capture, not this request.
      ...(persistence
        ? {
            persistence: {
              load: () => persistence.load(),
              save: (t: BearerTokens) => {
                try {
                  persistence.save(t);
                } catch (e) {
                  reportCacheWriteFailure(e);
                }
              },
              clear: () => {
                try {
                  persistence.clear();
                } catch {
                  /* nothing useful to do; the next capture overwrites it */
                }
              },
            },
          }
        : {}),
    });
    return this.#tokens;
  }

  #client(): ApiClient {
    if (this.#api) return this.#api;
    this.#api = createApiClient({
      baseUrl: this.#baseUrl,
      tokenManager: this.#tokenManager(),
      serviceName: 'Outlook',
      fetchImpl: this.#fetchImpl,
      timeout: 60_000,
      baseHeaders: { Accept: 'application/json' },
      onUnauthorized: () =>
        new McpToolError('Outlook rejected the access token (401).', {
          hint:
            'The token has expired or been revoked. Open a signed-in Outlook tab ' +
            'and retry — the next call re-captures automatically.',
        }),
    });
    return this.#api;
  }

  /**
   * GET a path relative to the API base.
   *
   * `text` asks Outlook to render message bodies as plain text instead of HTML
   * — measured 9.4x smaller on a real message, and it also normalises a mixed
   * HTML/Text collection to one type. `prefer` sets the `Prefer` header
   * directly, for anything else the API takes there (notably
   * `outlook.timezone="…"`).
   */
  async get<T>(path: string, opts: { text?: boolean; prefer?: string } = {}): Promise<T> {
    this.#requireConfigured();
    const prefer = [
      opts.text ? 'outlook.body-content-type="text"' : undefined,
      opts.prefer,
    ].filter((p): p is string => p !== undefined);
    return this.#client().fetchJson<T>('GET', path, {
      headers: prefer.length ? { Prefer: prefer.join(', ') } : undefined,
    });
  }

  /** GET an absolute URL — used only to follow `@odata.nextLink`. */
  async getAbsolute<T>(url: string): Promise<T> {
    this.#requireConfigured();
    const base = new URL(this.#baseUrl);
    const target = new URL(url);
    if (target.origin !== base.origin) {
      throw new McpToolError(`Refusing to follow a link to ${target.origin}.`, {
        hint: 'Pagination links must stay on the Outlook API host.',
      });
    }
    return this.#client().fetchJson<T>('GET', target.pathname + target.search);
  }

  /**
   * The single write path. Everything mutating goes through here so auth and
   * error shaping stay in one place.
   */
  async write<T>(
    method: 'POST' | 'PATCH' | 'DELETE',
    path: string,
    body?: unknown,
  ): Promise<T> {
    this.#requireConfigured();
    return this.#client().fetchJson<T>(method, path, {
      body: body === undefined ? undefined : body,
    });
  }

  /** The API base this client talks to. Non-secret; safe for healthcheck output. */
  get apiBase(): string {
    return this.#baseUrl;
  }

  /**
   * Ensure a usable token is held, capturing or refreshing only if needed.
   *
   * `getAccessToken` is the same call every request makes, so this exercises
   * the real resolution path rather than a re-derivation of it — the thing that
   * makes a passing healthcheck mean real tools work.
   */
  async refreshTokenIfNeeded(): Promise<void> {
    this.#requireConfigured();
    await this.#tokenManager().getAccessToken();
  }

  /** Force a fresh capture, ignoring the cached token. */
  async refreshToken(): Promise<void> {
    this.#requireConfigured();
    await this.#tokenManager().refreshNow();
  }

  /** Epoch ms at which the current token expires, or `null` if none is held. */
  tokenExpiresAt(): number | null {
    return this.#tokens ? this.#tokens.getExpiresAt() : null;
  }
}
