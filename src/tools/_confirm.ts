import type { CallToolResult } from '@modelcontextprotocol/server';
import { schemaConfirm, minifiedResult } from '@chrischall/mcp-utils';

export { schemaConfirm };

/**
 * Confirm-gate for a mutating tool (the fleet convention). Without
 * `confirm: true` the tool makes NO network call and returns a dry-run preview
 * of exactly what would be sent.
 *
 * This matters more here than in most repos: these writes send mail from the
 * user's real work address and move items in a live mailbox. A single
 * hallucinated call must not fire silently.
 *
 * Note this is a parameter, not an MCP elicitation. claude.ai declares no
 * elicitation capability and answers `-32021`, so a write that tried to ask
 * would simply fail there.
 */
export function previewUnlessConfirmed(
  confirm: boolean | undefined,
  action: string,
  method: string,
  path: string,
  body?: unknown,
): CallToolResult | null {
  if (confirm === true) return null;
  return minifiedResult({
    dryRun: true,
    action,
    method,
    path,
    ...(body !== undefined ? { willSend: body } : {}),
    note: 'Re-run with confirm: true to execute.',
  });
}
