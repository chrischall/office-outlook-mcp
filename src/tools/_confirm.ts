import type { CallToolResult, InputRequiredResult, ServerContext } from '@modelcontextprotocol/server';
import {
  confirmationFromEnv,
  confirmTokenParam,
  requireConfirmationWithFallback,
} from '@chrischall/mcp-utils';

export { confirmTokenParam };

/** Appended to every gated tool's description. */
export const CONFIRM_DESCRIPTION =
  ' Asks the user to confirm first: a confirmation prompt where the client supports one; otherwise ' +
  'the first call returns a preview and a confirmToken, and only a repeat call with that token ' +
  'proceeds (see MCP_CONFIRM_MODE).';

export interface WriteConfirmation {
  /** The tool name the token is bound to. */
  tool: string;
  /** Stable `<service>.<verb>` identifier. */
  action: string;
  /** The prompt shown before the preview. */
  message: string;
  /** Human-readable summary of what the write does. */
  summary: string;
  /** The primary id acted on, or '' when the write creates something new. */
  target: string;
  method: string;
  path: string;
  /** EXACTLY what the write will send. */
  body: unknown;
  /** The phase-2 token from the tool's input, if any. */
  confirmToken: string | undefined;
}

/**
 * Confirm-gate for a mutating tool. Returns `undefined` when the write may
 * proceed; anything else is the result to hand back unchanged, and in that
 * case NO network write has been made.
 *
 * This matters more here than in most repos: these writes send mail from the
 * user's real work address and move items in a live mailbox. A single
 * hallucinated call must not fire silently. A client that can show a prompt
 * gets one; one that cannot (claude.ai declares no elicitation capability)
 * gets the two-phase token flow configured by `MCP_CONFIRM_MODE`.
 *
 * The token binds the method, path and body, so a token minted for one write
 * never authorises a different one.
 */
export function confirmWrite(
  ctx: ServerContext,
  w: WriteConfirmation,
): Promise<InputRequiredResult | CallToolResult | undefined> {
  const preview = { action: w.summary, method: w.method, path: w.path, willSend: w.body };
  return requireConfirmationWithFallback(
    ctx,
    confirmationFromEnv({
      action: w.action,
      message: w.message,
      details: preview,
      tool: w.tool,
      confirmToken: w.confirmToken,
      subject: () => ({
        target: w.target,
        payload: { method: w.method, path: w.path, body: w.body },
        preview,
      }),
    }),
  );
}
