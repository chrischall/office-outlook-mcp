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
 *   - older tenants still serve the app from `outlook.office.com`; those set
 *     `OUTLOOK_WEB_HOST=outlook.office.com`. Only ONE host is ever declared —
 *     see `captureHost` for why.
 *
 * We do NOT try to mint a token by calling an endpoint ourselves. There is no
 * such endpoint reachable from the page's origin — the token comes from MSAL's
 * hidden-iframe flow against login.microsoftonline.com, which a bridge fetch
 * cannot drive. Capture is the only path, which is why a failed capture is a
 * hard error with actionable guidance rather than a fallback.
 */
import { captureBearerHeader, createFetchproxyTransport, type CaptureTransport } from '@chrischall/mcp-utils/fetchproxy';
import { readEnvVar, readPortEnv, readTtlMsEnv } from '@chrischall/mcp-utils';
import { PACKAGE_NAME, VERSION } from './version.js';

export { stripBearer } from '@chrischall/mcp-utils/fetchproxy';

/**
 * The fetchproxy concentrator port.
 *
 * ONE port for the whole fleet: the ContextMint Bridge extension dials this port and
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
export function getWsPort(env: NodeJS.ProcessEnv = process.env): number {
  return readPortEnv('OUTLOOK_WS_PORT', DEFAULT_WS_PORT, { env });
}

/**
 * The two Outlook Web origins a signed-in tab can be on: the new one first,
 * then the legacy one still serving tenants not yet migrated.
 *
 * Only one is ever DECLARED to the bridge (`captureHost`). Both stay listed
 * for validating `OUTLOOK_WEB_HOST` and for a hosted registration's egress.
 */
export const OUTLOOK_WEB_HOSTS = ['outlook.cloud.microsoft', 'outlook.office.com'] as const;
export type OutlookWebHost = (typeof OUTLOOK_WEB_HOSTS)[number];

/**
 * The ONE host to capture from and declare to the bridge: `OUTLOOK_WEB_HOST`,
 * or `outlook.cloud.microsoft` by default.
 *
 * This declared both hosts until 0.8.1, and the bridge opens a background
 * relay tab for every declared domain that has no tab open, after every hello.
 * A signed-in `outlook.office.com` tab redirects to `outlook.cloud.microsoft`,
 * so no tab ever matched the second domain and each capture opened another
 * one. The second capture also ran after the first (the bridge serializes
 * them), so a capture from the legacy host waited out the other host's window
 * first.
 *
 * Declaring one host fixes both. It is also the smallest scope the bridge can
 * be asked for, and the same on every hello, so it never offers a grant again.
 * Changing the host changes the domain set, which the bridge re-pairs on — a
 * one-time approval.
 *
 * An unrecognised value throws rather than falling back: capturing from the
 * wrong host just times out, with nothing saying why.
 */
export function captureHost(env: NodeJS.ProcessEnv = process.env): OutlookWebHost {
  const raw = readEnvVar('OUTLOOK_WEB_HOST', { env });
  if (raw === undefined) return OUTLOOK_WEB_HOSTS[0];
  const host = raw.trim().toLowerCase();
  const known = OUTLOOK_WEB_HOSTS.find((h) => h === host);
  if (known === undefined) {
    throw new Error(
      `OUTLOOK_WEB_HOST=${JSON.stringify(raw)} is not an Outlook Web host. Use ` +
        `${OUTLOOK_WEB_HOSTS.join(' or ')} — whichever your signed-in Outlook tab is on — or unset it.`,
    );
  }
  return known;
}

/**
 * How long to wait for the page to make a request we can read.
 *
 * The window IS the mechanism: capture resolves on the NEXT matching request,
 * so nothing the page did before the listener came up counts. Outlook Web polls
 * on its own, but an idle background tab can stay quiet for a while — 30s is
 * long enough that ordinary polling or a user glance lands inside it, short
 * enough that an idle tab does not hang a tool call for minutes.
 */
const DEFAULT_CAPTURE_WINDOW_MS = 30_000;

/**
 * Build the capture transport through this module's import binding, so the
 * test can stand in for the bridge. mcp-utils' own default calls its internal
 * reference, which a module mock cannot reach.
 */
const createCaptureTransport = (opts: Parameters<typeof createFetchproxyTransport>[0]): CaptureTransport =>
  createFetchproxyTransport<CaptureTransport>(opts);

/**
 * Capture an Outlook access token from the user's signed-in browser tab on
 * `captureHost(env)`. Returns the bare token (no `Bearer ` prefix).
 */
export async function captureTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  const host = captureHost(env);
  // `captureBearerHeader` declares only `capture_request_header` beside the
  // `fetch` base verb, derives the bridge deadline from the window (window +
  // 15 s, never under 30 s) so a raised window is not cut off early, starts the
  // transport (skipping `start()` makes the first verb throw "ensureConnected
  // called before listen()", which reads like a pairing problem), rejects a
  // header too short to be a token (real ones are ~5 KB of JWT), and closes
  // the transport whatever happened.
  return captureBearerHeader({
    captures: [{ host, path: '/*', headerName: 'authorization' }],
    // Exactly the host captured from: every API call afterwards is a plain
    // server-side `fetch` the bridge never sees. (The apex `cloud.microsoft` /
    // `office.com` once declared here granted the whole of Microsoft 365.)
    domains: [host],
    // Read per capture, not at import, so the env this server was given wins.
    windowMs: readTtlMsEnv('OUTLOOK_CAPTURE_TIMEOUT', DEFAULT_CAPTURE_WINDOW_MS, { env }),
    port: getWsPort(env),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: 'an Outlook token',
    guidance:
      `Open a signed-in Outlook tab on ${host}, leave it in the foreground for a ` +
      'moment so the page makes a request, and retry. If your Outlook is on the ' +
      `other host (${OUTLOOK_WEB_HOSTS.filter((h) => h !== host).join(', ')}), set ` +
      'OUTLOOK_WEB_HOST to it.',
    createTransport: createCaptureTransport,
  });
}
