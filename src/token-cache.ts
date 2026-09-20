import {
  createFileStatePersistence,
  resolveStateFile,
  type BearerTokens,
  type SyncStatePersistence,
} from '@chrischall/mcp-utils/session';
import { readEnvVar, parseBoolEnv } from '@chrischall/mcp-utils';

/**
 * Where the captured Outlook access token is cached between runs.
 *
 * Deliberately its OWN directory. The `outlook-fpx` access skill keeps a token
 * in `~/.outlook-fpx/curlrc` in a different format; writing there would corrupt
 * whichever of the two ran second.
 */
export function tokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'OUTLOOK_TOKEN_FILE',
    subdir: '.office-outlook-mcp',
    fileName: 'token.json',
  });
}

function isToken(raw: unknown): raw is BearerTokens {
  if (raw === null || typeof raw !== 'object') return false;
  const t = raw as Partial<BearerTokens>;
  return (
    typeof t.accessToken === 'string' &&
    t.accessToken !== '' &&
    typeof t.expiresAt === 'number' &&
    (t.refreshToken === undefined || typeof t.refreshToken === 'string')
  );
}

/**
 * What a cached token is bound to, or `null` when there is nothing worth
 * caching.
 *
 * An `OUTLOOK_ACCESS_TOKEN` IS the environment variable — there is no capture
 * to skip, so caching it would only copy a credential onto disk. A captured
 * token is worth caching most of all: it lets a cold start proceed with no
 * browser present, which on a host that has none is the difference between
 * working and not.
 */
function bindingFor(env: NodeJS.ProcessEnv): string | null {
  if (readEnvVar('OUTLOOK_ACCESS_TOKEN', { env }) !== undefined) return null;
  if (parseBoolEnv('OUTLOOK_DISABLE_FETCHPROXY', { env })) return null;
  return 'fetchproxy';
}

/** The token cache, or `null` when it is off or there is nothing to cache. */
export function createTokenCache(
  env: NodeJS.ProcessEnv = process.env,
): SyncStatePersistence<BearerTokens> | null {
  if (!parseBoolEnv('OUTLOOK_TOKEN_CACHE', { env, default: true })) return null;
  const boundTo = bindingFor(env);
  if (boundTo === null) return null;

  return createFileStatePersistence<BearerTokens>({
    filePath: tokenCachePath(env),
    boundTo,
    validate: (raw) => (isToken(raw) ? raw : null),
  });
}

/**
 * Report a cache write that failed. Not fatal — the token is re-capturable —
 * but worth saying, because a read-only data dir otherwise looks exactly like a
 * server that simply never caches.
 *
 * stderr only; stdout is the JSON-RPC channel.
 */
export function reportCacheWriteFailure(err: unknown): void {
  const detail = err instanceof Error ? err.message : String(err);
  console.error(
    `[office-outlook-mcp] could not cache the access token (${detail}); continuing ` +
      'without the cache — every restart will re-capture from the browser.',
  );
}
