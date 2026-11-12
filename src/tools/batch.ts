import type { McpServer } from '@modelcontextprotocol/server';
import { registerBatchTool, registerConfirmTool } from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { workMeta } from './_hints.js';
import { withToolMeta } from './_write.js';

/** The most ops one outlook_batch takes. */
export const BATCH_MAX_OPS = 100;

/**
 * `outlook_confirm` and `outlook_batch` (mcp-utils, docs/CONFIRM-TOKEN.md
 * §13–14):
 *
 * - `outlook_confirm({confirmToken})` finishes any gated write's two-step
 *   confirmation with the token alone: the server kept the phase-1 arguments.
 * - `outlook_batch({ops})` previews and approves several mailbox writes at
 *   once — create a folder, then a rule filing into it with
 *   refs {"moveToFolder": "0.id"} — in one approval.
 *
 * Only the writes registered with `registerWrite` are batchable: mailbox
 * changes (moves, read state, flags, rules, folders). Nothing that emails
 * other people — sends, replies, invites, events, drafts, auto-replies — can
 * ride in a batch.
 *
 * Registered LAST: the batch tool's description lists the writes registered
 * before it.
 */
export function registerBatchTools(server: McpServer, _client: OutlookClient): void {
  withToolMeta(server, workMeta('other'), () => registerConfirmTool(server, { name: 'outlook_confirm' }));
  withToolMeta(server, workMeta('change', { aliases: ['several changes at once', 'batch', 'bulk changes'] }), () =>
    registerBatchTool(server, {
      name: 'outlook_batch',
      maxOps: BATCH_MAX_OPS,
      description:
        'Several mailbox writes the user approves at once, e.g. outlook_create_folder then outlook_create_rule filing into it with refs {"moveToFolder": "0.id"}, or a handful of moves and flags. There is no rollback: if an op fails, the ones before it stay applied and the rest do not run. A write Outlook did not confirm (a timeout or server error after sending) is not a failure here: its result carries status "unknown" and the batch goes on, so check each result before relying on it. To move MANY messages matching a condition (a sender, an age, read state), use outlook_move_messages instead — one call, one confirmation for the whole set.',
    }),
  );
}
