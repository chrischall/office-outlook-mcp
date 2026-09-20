import type { McpServer } from '@modelcontextprotocol/server';
import { registerCredentialHealthcheckTool } from '@chrischall/mcp-utils/healthcheck';
import type { OutlookClient } from '../client.js';

/**
 * The CREDENTIAL healthcheck, not the bridge one.
 *
 * Classify by where the transport is HELD, not by what is imported: this repo
 * constructs a fetchproxy transport ephemerally inside the token capture and
 * closes it immediately, then talks to `outlook.office.com` over plain fetch.
 * The bridge is not on the request path, so bridge role/port would be the wrong
 * thing to report — what a broken install needs to know is whether a token
 * resolved and whether Outlook accepted it.
 *
 * Imported from `/healthcheck`, deliberately NOT from `/fetchproxy`: that
 * module pulls in the optional `@fetchproxy/server` peer, which is absent in
 * the `.mcpb` bundle.
 */
export function registerHealthcheckTool(server: McpServer, client: OutlookClient): void {
  registerCredentialHealthcheckTool({
    server,
    prefix: 'outlook',
    hostLabel: 'outlook.office.com',
    probePath: '/api/v2.0/me',
    /**
     * Resolve the credential the way the real tools do — same TokenManager,
     * same cache, same capture — so a passing healthcheck means real tools
     * work. Returns a SOURCE LABEL and non-secret detail, never the token:
     * this result is what people paste into a chat when something is broken.
     */
    resolveCredential: async () => {
      await client.refreshTokenIfNeeded();
      const expiresAt = client.tokenExpiresAt();
      return {
        source: client.tokenSource,
        detail: {
          expiresAt: expiresAt === null ? null : new Date(expiresAt).toISOString(),
          minutesRemaining:
            expiresAt === null ? null : Math.round((expiresAt - Date.now()) / 60_000),
          apiBase: client.apiBase,
        },
      };
    },
    probeFn: async () => {
      const me = await client.get<{ EmailAddress?: string; DisplayName?: string }>('/me');
      return `authenticated as ${me?.EmailAddress ?? 'unknown'}`;
    },
  });
}
