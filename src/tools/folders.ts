import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  confirmTokenParam,
  confirmWrite,
  errorStatusOf,
  McpToolError,
  minifiedResult,
} from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { compactFolder, type OutlookFolder } from '../view.js';
import { workMeta } from './_hints.js';
import { WELL_KNOWN } from './mail.js';
import { folderId, readBack, unverified } from './rules.js';
import { TIMEOUT_SENTENCE, writeOrUnknown } from './writes.js';

/** The mailbox's top level: the parent of Inbox, Archive and the folders beside them. */
const ROOT = 'msgfolderroot';
const RULES_PATH = '/me/MailFolders/inbox/MessageRules';
const folderPath = (id: string) => `/me/mailfolders/${encodeURIComponent(id)}`;

const PROTECTED = new Set<string>([...WELL_KNOWN, ROOT]);
const LIST_HINT = 'Pass a folder id from outlook_list_folders or a well-known name such as "inbox".';

/** Outlook's own folders cannot be renamed, moved or deleted, so refuse them by name before any request. */
function refuseWellKnown(id: string, verb: string): void {
  if (PROTECTED.has(id.toLowerCase())) {
    throw new McpToolError(`"${id}" is one of Outlook's own folders and cannot be ${verb}.`, {
      hint: 'Pick a folder you created; list them with outlook_list_folders.',
    });
  }
}

async function readFolder(client: OutlookClient, id: string): Promise<OutlookFolder> {
  try {
    return await client.get<OutlookFolder>(folderPath(id));
  } catch (e) {
    const status = errorStatusOf(e);
    if (status !== 400 && status !== 404) throw e;
    throw new McpToolError(`No mail folder matches "${id}".`, { hint: LIST_HINT });
  }
}

const shown = (f: OutlookFolder) => ({ ...compactFolder(f), Parent: f.ParentFolderId });

/**
 * Outlook deletes a big folder's mail after the DELETE is accepted, at about 75
 * messages a second (live 2026-11-10: 13,370 took ~3 minutes, 63,466 ~14), so
 * the call outlives the client timeout. A resend while it runs would only
 * restart a delete that is already happening.
 */
const STILL_DELETING =
  'Outlook did not answer in time. A large folder keeps running its delete on the server for minutes after the request (about 75 messages a second), so it is probably still deleting. Check the folder\'s Total in outlook_list_folders: falling means it is working, gone means done. Do not resend the delete.';

const plural = (n: number, one: string) => `${n} ${one}${n === 1 ? '' : 's'}`;

/** What goes with a folder when it is deleted, in words. */
function contents(f: OutlookFolder): string {
  const total = f.TotalItemCount ?? 0;
  const unread = f.UnreadItemCount ?? 0;
  const parts = [`${plural(total, 'message')}${unread ? ` (${unread} unread)` : ''}`];
  if (f.ChildFolderCount) parts.push(`${plural(f.ChildFolderCount, 'subfolder')} and everything in them`);
  return parts.join(' and ');
}

type Rule = { DisplayName?: string; Actions?: { MoveToFolder?: string; CopyToFolder?: string } };

/**
 * Rules that file mail into this folder. Losing the folder breaks them, so the
 * confirmation names them. A failed read only loses the warning, not the delete.
 */
async function rulesUsing(client: OutlookClient, id: string): Promise<string[]> {
  try {
    const page = await client.get<{ value?: Rule[] }>(RULES_PATH);
    return (page.value ?? [])
      .filter((r) => r.Actions?.MoveToFolder === id || r.Actions?.CopyToFolder === id)
      .map((r) => r.DisplayName ?? '(unnamed rule)');
  } catch {
    return [];
  }
}

export function registerFolderTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_create_folder',
    {
      description:
        'Create a mail folder: at the top level of the mailbox by default, or inside `parent`. Nothing reaches anyone else and an empty folder is easily deleted, so this does not ask for confirmation. Point rules or moves at the returned Id. ' +
        TIMEOUT_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      _meta: workMeta('change', { aliases: ['new folder', 'add folder', 'make folder'] }),
      inputSchema: z.object({
        name: z.string().min(1).describe('Folder name'),
        parent: z
          .string()
          .min(1)
          .optional()
          .describe('Folder id or well-known name (e.g. "inbox") to create it inside. Omit for the top level.'),
      }),
    },
    async ({ name, parent }) => {
      const path = `${folderPath(parent ?? ROOT)}/childfolders`;
      let write;
      try {
        write = await writeOrUnknown(() => client.write<OutlookFolder>('POST', path, { DisplayName: name }), 'outlook_list_folders');
      } catch (e) {
        const status = errorStatusOf(e);
        if (status === 409) {
          throw new McpToolError(`A folder named "${name}" already exists there.`, {
            hint: 'Use the existing one (see outlook_list_folders) or pick another name.',
          });
        }
        if (parent !== undefined && (status === 400 || status === 404)) {
          throw new McpToolError(`No mail folder matches "${parent}".`, { hint: LIST_HINT });
        }
        throw e;
      }
      if (!write.ok) return write.result;
      return minifiedResult({ created: true, ...shown(write.value ?? { DisplayName: name }) });
    },
  );

  server.registerTool(
    'outlook_update_folder',
    {
      description:
        'Rename a mail folder, move it under another folder, or both. Its mail and subfolders go with it, and rules that file into it keep working. Outlook\'s own folders (Inbox, Archive, …) are refused. The result is verified by re-reading the folder. ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      _meta: workMeta('change', { aliases: ['rename folder', 'move folder'] }),
      inputSchema: z.object({
        id: z.string().min(1).describe('Folder Id from outlook_list_folders'),
        name: z.string().min(1).optional().describe('New name'),
        parent: z
          .string()
          .min(1)
          .optional()
          .describe('Folder id or well-known name to move it into; "msgfolderroot" is the top level'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, name, parent, confirmToken }, ctx) => {
      if (name === undefined && parent === undefined) {
        throw new McpToolError('Nothing to change on this folder.', { hint: 'Pass `name`, `parent` or both.' });
      }
      refuseWellKnown(id, 'renamed or moved');
      if (parent !== undefined && parent === id) {
        throw new McpToolError('A folder cannot be moved into itself.', { hint: 'Pick a different `parent`.' });
      }
      const current = await readFolder(client, id);
      const destination = parent === undefined ? undefined : await folderId(client, parent);
      const label = current.DisplayName ?? id;
      const changes = [
        ...(name !== undefined ? [`rename it to "${name}"`] : []),
        ...(parent !== undefined ? [parent.toLowerCase() === ROOT ? 'move it to the top level' : `move it into ${parent}`] : []),
      ].join(' and ');
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_update_folder',
        action: 'mail.folder_update',
        message: `Review and confirm: ${changes}.`,
        summary: `Folder "${label}": ${changes}.`,
        account: undefined,
        target: id,
        // A rename and a move are separate calls; preview both, as sent.
        payload: [
          ...(name !== undefined ? [{ method: 'PATCH', path: folderPath(id), body: { DisplayName: name } }] : []),
          ...(destination !== undefined
            ? [{ method: 'POST', path: `${folderPath(id)}/move`, body: { DestinationId: destination } }]
            : []),
        ],
        confirmToken,
        preview: { current: shown(current) },
      });
      if (gate) return gate;

      if (name !== undefined) {
        const rename = await writeOrUnknown(() => client.write('PATCH', folderPath(id), { DisplayName: name }), 'outlook_list_folders');
        if (!rename.ok) return rename.result;
      }
      let finalId = id;
      if (destination !== undefined) {
        const move = await writeOrUnknown(
          () => client.write<OutlookFolder>('POST', `${folderPath(id)}/move`, { DestinationId: destination }),
          'outlook_list_folders',
        );
        if (!move.ok) return move.result;
        finalId = move.value?.Id ?? id;
      }

      const read = await readBack(() => client.get<OutlookFolder>(folderPath(finalId)));
      if (!read.ok) return minifiedResult({ updated: true, Id: finalId, ...unverified('outlook_list_folders') });
      const after = read.value;
      const stale = [
        ...(name !== undefined && after.DisplayName !== name ? ['name'] : []),
        ...(destination !== undefined && after.ParentFolderId !== destination ? ['location'] : []),
      ];
      return minifiedResult({
        updated: stale.length === 0,
        ...shown(after),
        ...(stale.length ? { warning: `Outlook accepted the write but the ${stale.join(' and ')} did not change.` } : {}),
      });
    },
  );

  server.registerTool(
    'outlook_delete_folder',
    {
      description:
        'Delete a mail folder together with every message and subfolder in it. This is permanent: the folder does NOT go to Deleted Items, so it cannot be restored from there. To keep the mail, move it out first (outlook_move_message). The confirmation says how much mail goes with it and names any inbox rules that file into it (those rules stop working). Outlook\'s own folders are refused. The result is verified by re-reading the folder. A folder holding thousands of messages takes minutes to delete, so the call can come back with status "unknown" while Outlook keeps deleting: check its Total in outlook_list_folders falling, and never resend. ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
      _meta: workMeta('change', { aliases: ['remove folder'] }),
      inputSchema: z.object({
        id: z.string().min(1).describe('Folder Id from outlook_list_folders'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, confirmToken }, ctx) => {
      refuseWellKnown(id, 'deleted');
      const current = await readFolder(client, id);
      const label = current.DisplayName ?? id;
      const rules = await rulesUsing(client, current.Id ?? id);
      const ruleNote = rules.length
        ? ` These inbox rules file mail into it and will stop working: ${rules.map((r) => `"${r}"`).join(', ')}.`
        : '';
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_delete_folder',
        action: 'mail.folder_delete',
        message: `Review and confirm permanently deleting this folder with its ${contents(current)} (it does not go to Deleted Items):${ruleNote}`,
        summary: `Permanently delete folder "${label}" with its ${contents(current)}; it does not go to Deleted Items.${ruleNote}`,
        account: undefined,
        target: id,
        request: { method: 'DELETE', path: folderPath(id) },
        confirmToken,
        preview: { current: shown(current), ...(rules.length ? { rulesUsingIt: rules } : {}) },
      });
      if (gate) return gate;

      const write = await writeOrUnknown(() => client.write('DELETE', folderPath(id)), 'outlook_list_folders');
      if (!write.ok) return minifiedResult({ status: 'unknown', Id: id, warning: STILL_DELETING });
      const read = await readBack(() => client.get<OutlookFolder>(folderPath(id)));
      if (!read.ok && errorStatusOf(read.error) === 404) return minifiedResult({ deleted: true, Id: id });
      if (!read.ok) return minifiedResult({ deleted: true, Id: id, ...unverified('outlook_list_folders') });
      return minifiedResult({ deleted: false, Id: id, warning: 'Outlook accepted the delete but the folder is still there.' });
    },
  );
}
