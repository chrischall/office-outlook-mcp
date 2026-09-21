/**
 * Bootstrap-only fetchproxy auth.
 *
 * Outlook Web holds its access token in JS memory and sends it as
 * `Authorization: Bearer …` on every API call. We snapshot that header off a
 * request the page already makes, and from then on this MCP talks to
 * `https://outlook.office.com/api/v2.0` with a plain Node `fetch`.
 *
 * The bridge is therefore touched ONCE per token (~25h), never in the hot path.
 * That is what makes this repo hostable: a full-bridge MCP needs a browser
 * attached for every call, this one needs it only to mint.
 *
 * Verified live 2026-09-20 against a real M365 mailbox:
 *   - the signed-in tab is on `outlook.cloud.microsoft` (the new Outlook Web
 *     domain), but the token it carries has `aud: https://outlook.office.com`;
 *   - that token authenticates plain server-side requests to
 *     `outlook.office.com/api/v2.0` with no bridge involved;
 *   - older tenants still serve the app from `outlook.office.com`, so both
 *     hosts are declared and captured concurrently.
 *
 * We do NOT try to mint a token by calling an endpoint ourselves. There is no
 * such endpoint reachable from the page's origin — the token comes from MSAL's
 * hidden-iframe flow against login.microsoftonline.com, which a bridge fetch
 * cannot drive. Capture is the only path, which is why a failed capture is a
 * hard error with actionable guidance rather than a fallback.
 */
import {
  createBootstrapOpts,
  createFetchproxyTransport,
  type Capability,
} from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv, readTtlMsEnv } from '@chrischall/mcp-utils';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The fetchproxy concentrator port.
 *
 * ONE port for the whole fleet: the Transporter extension dials this port and
 * servers host/peer-elect on it, so a "unique" per-MCP default would simply
 * never be found. `OUTLOOK_WS_PORT` overrides it for local development and —
 * the reason it exists — for a hosted bridged registration, where mcp-host
 * names this variable in `bridgePortEnv` and puts the child's port in it.
 *
 * `readPortEnv` rather than `Number(...)` because the value arrives from a host
 * template: an unsubstituted `${OUTLOOK_WS_PORT}` falls back to the default
 * instead of handing `NaN` to the server.
 */
const DEFAULT_WS_PORT = 37_149;
export function getWsPort(): number {
  return readPortEnv('OUTLOOK_WS_PORT', DEFAULT_WS_PORT);
}

/** New Outlook Web. This is where a current tenant's tab actually lives. */
const CAPTURE_DECL_CLOUD = {
  host: 'outlook.cloud.microsoft',
  path: '/*',
  headerName: 'authorization',
} as const;

/** Legacy host, still serving the app for tenants not yet migrated. */
const CAPTURE_DECL_OFFICE = {
  host: 'outlook.office.com',
  path: '/*',
  headerName: 'authorization',
} as const;

/**
 * How long to wait for the page to make a request we can read.
 *
 * The window IS the mechanism: capture resolves on the NEXT matching request,
 * so nothing the page did before the listener came up counts. Outlook Web polls
 * on its own, but an idle background tab can stay quiet for a while — 30s is
 * long enough that ordinary polling or a user glance lands inside it, short
 * enough that an idle tab does not hang a tool call for minutes.
 */
const CAPTURE_TIMEOUT_MS = readTtlMsEnv('OUTLOOK_CAPTURE_TIMEOUT', 30_000);

/**
 * The transport-wide deadline, derived from the capture window rather than left
 * at the library default.
 *
 * `captureRequestHeader`'s own `timeoutMs` is silently capped by
 * `fetchTimeoutMs`: a window raised past the deadline fails early at the
 * deadline's number, which reads as a timeout nobody configured. Deriving it
 * means the two cannot disagree.
 */
const BRIDGE_DEADLINE_MS = Math.max(CAPTURE_TIMEOUT_MS + 15_000, 30_000);

/**
 * Shortest plausible token. Real ones are ~5 KB of JWT; this only guards
 * against an empty or echoed header, not against a wrong-audience token.
 */
const MIN_TOKEN_LENGTH = 100;

/**
 * Add `fetch` to the derived capability set.
 *
 * `createBootstrapOpts` infers capabilities from the bootstrap DECLARATIONS and
 * the field REPLACES the server's defaults rather than extending them, so a
 * declaration-only set can silently lose `fetch`.
 */
function withFetch<T extends { capabilities?: readonly Capability[] }>(opts: T): T {
  const derived = opts.capabilities ?? [];
  return {
    ...opts,
    capabilities: derived.includes('fetch' as Capability)
      ? derived
      : ([...derived, 'fetch' as Capability] as readonly Capability[]),
  };
}

/**
 * Settle a set of per-host capture attempts under ONE overall deadline.
 *
 * `Promise.any` alone is not enough. The two declared hosts are documented as
 * raced, but the bridge serializes them behind its single connection, so the
 * wait was the sum rather than the max — measured live at 14.4s for a 5s window
 * and 36s for a 20s one. At the 30s default that lands past the 60s request
 * timeout an MCP client uses by default, so the host gave up before this code
 * could report the real reason. The window is a promise about total wait, and
 * this is what keeps it.
 *
 * Exported for the test that pins that bound.
 */
export async function raceCaptures(
  attempts: readonly Promise<string>[],
  windowMs: number,
): Promise<string> {
  // A rejection that loses the race must not surface as an unhandled rejection
  // once the winner has already settled the caller.
  for (const a of attempts) a.catch(() => {});

  let timer: ReturnType<typeof setTimeout> | undefined;
  const expiry = new Promise<never>((_, reject) => {
    timer = setTimeout(
      () =>
        reject(
          new Error(
            `timed out after ${Math.round(windowMs / 1000)}s waiting for the page to ` +
              'make a request we could read',
          ),
        ),
      windowMs,
    );
  });

  try {
    // `any` resolves on the first SUCCESS rather than the first settle, so the
    // host the user is not on failing does not sink the one they are on.
    return await Promise.race([
      Promise.any(attempts).catch((e: unknown) => {
        throw new Error(
          e instanceof AggregateError
            ? e.errors.map((x) => (x as Error).message).join('; ')
            : (e as Error).message,
        );
      }),
      expiry,
    ]);
  } finally {
    clearTimeout(timer);
  }
}

/** Normalise a captured header to a bare token, dropping any `Bearer ` prefix. */
export function stripBearer(raw: string): string {
  const m = /^\s*Bearer\s+(.+)$/i.exec(raw);
  return (m ? m[1] : raw).trim();
}

/**
 * Capture an Outlook access token from the user's signed-in browser tab.
 *
 * Both hosts are raced: whichever the tab is actually on answers first, and the
 * other simply never resolves. Returns the bare token (no `Bearer ` prefix).
 */
export async function captureTokenViaFetchproxy(): Promise<string> {
  const transport = createFetchproxyTransport({
    ...withFetch(
      createBootstrapOpts({
        domains: ['cloud.microsoft', 'office.com'],
        bootstrap: {
          captureHeaders: [{ ...CAPTURE_DECL_CLOUD }, { ...CAPTURE_DECL_OFFICE }],
        },
      }),
    ),
    port: getWsPort(),
    fetchTimeoutMs: BRIDGE_DEADLINE_MS,
    serverName: PACKAGE_NAME,
    version: VERSION,
  });

  try {
    // `start()` loads the identity keypair. It does not bind the port or dial —
    // that stays lazy until the first verb — but skipping it makes the first
    // verb throw "ensureConnected called before listen()", which reads exactly
    // like a pairing problem and is not one.
    await transport.start();

    const attempts = [CAPTURE_DECL_CLOUD, CAPTURE_DECL_OFFICE].map((decl) =>
      transport.server
        .captureRequestHeader({ ...decl, timeoutMs: CAPTURE_TIMEOUT_MS })
        .then((raw) => {
          if (typeof raw !== 'string') {
            throw new Error(`${decl.host}: capture returned no header`);
          }
          const token = stripBearer(raw);
          if (token.length < MIN_TOKEN_LENGTH) {
            throw new Error(
              `${decl.host}: authorization header was present but too short to be a ` +
                `token (${token.length} chars)`,
            );
          }
          return token;
        }),
    );

    try {
      return await raceCaptures(attempts, CAPTURE_TIMEOUT_MS);
    } catch (e) {
      const detail = (e as Error).message;
      throw new Error(
        `could not capture an Outlook token from the browser (${detail}). ` +
          'Open a signed-in Outlook tab (outlook.cloud.microsoft or ' +
          'outlook.office.com), leave it in the foreground for a moment so the ' +
          'page makes a request, and retry.',
      );
    }
  } finally {
    try {
      await transport.close();
    } catch {
      /* the token, or the original failure, is what matters */
    }
  }
}
