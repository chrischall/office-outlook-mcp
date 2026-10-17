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
 *     hosts are declared and captured, under ONE overall window.
 *
 * We do NOT try to mint a token by calling an endpoint ourselves. There is no
 * such endpoint reachable from the page's origin — the token comes from MSAL's
 * hidden-iframe flow against login.microsoftonline.com, which a bridge fetch
 * cannot drive. Capture is the only path, which is why a failed capture is a
 * hard error with actionable guidance rather than a fallback.
 */
import { captureBearerHeader } from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv, readTtlMsEnv } from '@chrischall/mcp-utils';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * `raceCaptures` settles the per-host attempts under ONE overall deadline:
 * `Promise.any` alone is not enough, because the bridge serializes the two
 * hosts behind its single connection, so the wait was the sum rather than the
 * max — measured live at 14.4s for a 5s window and 36s for a 20s one. At the
 * 30s default that lands past the 60s request timeout an MCP client uses by
 * default. Re-exported (it is mcp-utils', lifted from here) for the test that
 * pins that bound.
 */
export { raceCaptures, stripBearer } from '@chrischall/mcp-utils/fetchproxy';

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
 * The hosts a token may be captured from — the two Outlook Web origins.
 *
 * Exported so the trust boundary below can be derived from them rather than
 * restated.
 */
export const CAPTURE_HOSTS = [CAPTURE_DECL_CLOUD.host, CAPTURE_DECL_OFFICE.host] as const;

/**
 * The extension's trust boundary: exactly the hosts we capture from.
 *
 * These were the apex domains `cloud.microsoft` and `office.com`, which the
 * extension reads as "any subdomain of either" — the whole of Microsoft 365
 * granted to read one header off two known hosts. It is also what sent the
 * pairing flow to `m365.cloud.microsoft/chat`, a tab nothing here needs.
 *
 * Narrowing costs nothing: the capture declarations were always these two
 * hosts, and every API call the server makes afterwards is a plain server-side
 * `fetch` that the bridge never sees. Verified live 2026-09-20 — capture still
 * succeeded, and an already-paired extension did not ask to re-pair.
 */
export const TRUST_DOMAINS = [...CAPTURE_HOSTS];

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
 * Capture an Outlook access token from the user's signed-in browser tab.
 *
 * Both hosts are attempted: whichever the tab is actually on answers first, and
 * the other simply never resolves. The bridge serializes them, so the pair is
 * held to one deadline by `raceCaptures` rather than to one deadline each.
 * Returns the bare token (no `Bearer ` prefix).
 */
export async function captureTokenViaFetchproxy(): Promise<string> {
  // `captureBearerHeader` declares only `capture_request_header` beside the
  // `fetch` base verb, derives the bridge deadline from the window (window +
  // 15 s, never under 30 s) so a raised window is not cut off early, starts the
  // transport (skipping `start()` makes the first verb throw "ensureConnected
  // called before listen()", which reads like a pairing problem), rejects a
  // header too short to be a token (real ones are ~5 KB of JWT), and closes
  // the transport whatever happened.
  return captureBearerHeader({
    captures: [CAPTURE_DECL_CLOUD, CAPTURE_DECL_OFFICE],
    domains: TRUST_DOMAINS,
    windowMs: CAPTURE_TIMEOUT_MS,
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: 'an Outlook token',
    guidance:
      'Open a signed-in Outlook tab (outlook.cloud.microsoft or ' +
      'outlook.office.com), leave it in the foreground for a moment so the ' +
      'page makes a request, and retry.',
  });
}
