/**
 * A Sydney (M365 Copilot) access token, via MSAL silent refresh.
 *
 * Adapted from the sibling project `opencode-m365-copilot` (MIT, same owner):
 * https://github.com/chrischall/opencode-copilot-plugin — src/auth.ts, which
 * reverse-engineered this flow. Trimmed here to the silent-acquire path this
 * server needs.
 *
 * The client id below is Microsoft's own Office web Copilot application. We did
 * not register it and no one can: the Sydney scopes are granted to no other
 * client, so a loopback redirect is rejected (`AADSTS50011`) and the
 * device-code grant demands a client secret only Microsoft holds
 * (`AADSTS7000218`). The one-time interactive sign-in therefore has to drive a
 * real browser — which is exactly what `opencode-m365 login` does. This module
 * is SILENT-ONLY: it reads the MSAL cache that login wrote and refreshes from
 * it. When the cache cannot produce a token it raises a credential error
 * telling the user to run that login, never launching a browser from inside the
 * server process.
 *
 * Unlike the Outlook mail token (captured fresh from the browser every ~25h),
 * the Sydney token lives ~77min and is obtained by nested-app-auth that keeps it
 * in memory — it is never in a header or storage the fetchproxy bridge could
 * read. MSAL's refresh token is the only durable, server-usable credential, so
 * this is the only route that yields a self-refreshing Copilot connection.
 */
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { McpToolError, readEnvVar } from '@chrischall/mcp-utils';

/** Write owner-only, creating the directory. Holds a refresh token, so 0600. */
function writePrivateFile(file: string, data: string): void {
  mkdirSync(dirname(file), { recursive: true });
  writeFileSync(file, data, { mode: 0o600 });
  // writeFileSync's mode only applies when creating; tighten an existing file too.
  if (process.platform !== 'win32') chmodSync(file, 0o600);
}

/** Microsoft's Office web Copilot client. Not ours, and not replaceable. */
export const COPILOT_CLIENT_ID = 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1';
export const COPILOT_AUTHORITY = 'https://login.microsoftonline.com/common';

/** Scopes for the chat backend. The resulting token's audience is Sydney. */
export const SYDNEY_SCOPES = [
  'https://substrate.office.com/sydney/M365Chat.Read',
  'https://substrate.office.com/sydney/sydney.readwrite',
];

/** Re-acquire this far ahead of expiry rather than racing the clock mid-turn. */
const EXPIRY_MARGIN_MS = 120_000;

/**
 * Where `opencode-m365 login` writes its MSAL cache by default. Reading the same
 * file means a user who already uses that CLI needs no second sign-in;
 * `OUTLOOK_COPILOT_MSAL_CACHE` overrides it (e.g. to a cache written with
 * `M365_CACHE_FILE` pointed elsewhere).
 */
export function copilotCacheFile(env: NodeJS.ProcessEnv = process.env): string {
  return (
    readEnvVar('OUTLOOK_COPILOT_MSAL_CACHE', { env }) ??
    join(homedir(), '.config', 'opencode-copilot', 'msal-cache.json')
  );
}

/** The one shape of MSAL this module depends on; the real client satisfies it. */
export interface MsalLike {
  getTokenCache(): { getAllAccounts(): Promise<unknown[]> };
  acquireTokenSilent(request: {
    account: unknown;
    scopes: string[];
    forceRefresh?: boolean;
  }): Promise<{ accessToken?: string; expiresOn?: Date | null } | null>;
}

export interface CopilotTokenClient {
  /** A Sydney chat token. Silent refresh; throws a credential error if none. */
  getToken(): Promise<string>;
  /** Forget the cached token so the next call re-acquires. */
  invalidate(): void;
}

export interface CopilotTokenClientOptions {
  /** Inject an MSAL client; defaults to a real one over the on-disk cache. */
  app?: MsalLike;
  cacheFile?: string;
  env?: NodeJS.ProcessEnv;
}

/** Signing-in is a user action we cannot take for them; say how. */
function authRequired(reason: string): McpToolError {
  return new McpToolError(`${reason}.`, {
    hint:
      'Sign in once with `opencode-m365 login` (from the opencode-m365-copilot ' +
      'package) so an MSAL cache exists, or set OUTLOOK_COPILOT_MSAL_CACHE to one ' +
      'already written by it.',
    kind: 'no_credential',
  });
}

/** Persist MSAL's token cache to disk, owner-only; a write failure is non-fatal. */
function cachePlugin(file: string): import('@azure/msal-node').ICachePlugin {
  return {
    async beforeCacheAccess(context) {
      try {
        context.tokenCache.deserialize(readFileSync(file, 'utf8'));
      } catch {
        // First run or a corrupt cache: MSAL starts empty and the next call
        // surfaces "no signed-in account", pointing the user at the login.
      }
    },
    async afterCacheAccess(context) {
      if (!context.cacheHasChanged) return;
      try {
        writePrivateFile(file, context.tokenCache.serialize());
      } catch {
        // A read-only data dir costs the next cold start a refresh, not this call.
      }
    },
  };
}

interface Cached {
  token: string;
  expiresAt: number;
}

export function createCopilotTokenClient(options: CopilotTokenClientOptions = {}): CopilotTokenClient {
  const env = options.env ?? process.env;
  const cacheFile = options.cacheFile ?? copilotCacheFile(env);
  let app: MsalLike | undefined = options.app;
  let cached: Cached | undefined;

  async function client(): Promise<MsalLike> {
    if (app) return app;
    // Imported lazily so merely booting the server does not pull MSAL in — a
    // server never asked to reach Copilot pays nothing for it.
    const { PublicClientApplication } = await import('@azure/msal-node');
    app = new PublicClientApplication({
      auth: { clientId: COPILOT_CLIENT_ID, authority: COPILOT_AUTHORITY },
      cache: { cachePlugin: cachePlugin(cacheFile) },
    }) as unknown as MsalLike;
    return app;
  }

  return {
    async getToken(): Promise<string> {
      if (cached && cached.expiresAt - EXPIRY_MARGIN_MS > Date.now()) return cached.token;
      const msal = await client();
      const accounts = await msal.getTokenCache().getAllAccounts();
      const account = accounts[0];
      if (!account) throw authRequired('No signed-in Microsoft 365 Copilot account');

      let result: { accessToken?: string; expiresOn?: Date | null } | null;
      try {
        result = await msal.acquireTokenSilent({ account, scopes: SYDNEY_SCOPES });
      } catch {
        // Silent refresh fails when the refresh token expired or was revoked, or
        // the tenant now demands interaction (a new conditional-access policy).
        throw authRequired('Could not refresh the Microsoft 365 Copilot token silently');
      }
      if (!result?.accessToken) throw authRequired('Microsoft 365 returned no Copilot access token');

      cached = {
        token: result.accessToken,
        expiresAt: result.expiresOn ? new Date(result.expiresOn).getTime() : Date.now() + 300_000,
      };
      return cached.token;
    },
    invalidate() {
      cached = undefined;
    },
  };
}
