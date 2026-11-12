import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  confirmTokenParam,
  confirmWrite,
  defineWrite,
  errorStatusOf,
  McpToolError,
  minifiedResult,
} from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { compactFolder, type OutlookFolder } from '../view.js';
import { workMeta } from './_hints.js';
import { WELL_KNOWN } from './mail.js';
import { folderId, readBack, unverified } from './rules.js';
import { TIMEOUT_SENTENCE, writeOrUnknown, writeOrUnknownData } from './writes.js';
import { registerGatedWrite, registerUngatedWrite } from './_write.js';

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

/** A folder's own counts plus every folder beneath it. */
type Subtree = { total: number; unread: number; subfolders: number; ids: string[]; partial: boolean };

/** Past this many folders the walk stops and the counts read "at least". */
const SUBTREE_CAP = 500;
const CHILD_SELECT = 'Id,TotalItemCount,UnreadItemCount,ChildFolderCount';

/**
 * A folder's own TotalItemCount leaves out its subfolders, and a delete takes
 * them all: live 2026-11-10 a parent read "0 messages" over 2,725. So the
 * confirmation walks the tree. A failed or capped walk is reported as a lower
 * bound rather than blocking the delete.
 */
async function subtree(client: OutlookClient, root: OutlookFolder & { Id: string }): Promise<Subtree> {
  const out: Subtree = {
    total: root.TotalItemCount ?? 0,
    unread: root.UnreadItemCount ?? 0,
    subfolders: 0,
    ids: [root.Id],
    partial: false,
  };
  const queue: Array<{ id: string; expected: number }> =
    (root.ChildFolderCount ?? 0) > 0 ? [{ id: root.Id, expected: root.ChildFolderCount ?? 0 }] : [];
  while (queue.length) {
    const { id: parent, expected } = queue.shift() as { id: string; expected: number };
    if (out.subfolders >= SUBTREE_CAP) {
      out.partial = true;
      break;
    }
    let page: { value?: OutlookFolder[] };
    try {
      page = await client.get(`${folderPath(parent)}/childfolders`, { query: { $select: CHILD_SELECT, $top: 100 } });
    } catch {
      out.partial = true;
      continue;
    }
    const kids = page.value ?? [];
    // Hidden subfolders count in ChildFolderCount but are not listed (live: Conversation History).
    if (kids.length >= 100 || kids.length < expected) out.partial = true;
    for (const k of kids) {
      out.subfolders++;
      out.total += k.TotalItemCount ?? 0;
      out.unread += k.UnreadItemCount ?? 0;
      if (k.Id) {
        out.ids.push(k.Id);
        if ((k.ChildFolderCount ?? 0) > 0) queue.push({ id: k.Id, expected: k.ChildFolderCount ?? 0 });
      }
    }
  }
  return out;
}

/** What goes with a folder when it is deleted, in words. */
function contents(t: Subtree): string {
  const atLeast = t.partial ? 'at least ' : '';
  const mail = `${atLeast}${plural(t.total, 'message')}${t.unread ? ` (${t.unread} unread)` : ''}`;
  return t.subfolders ? `${mail}, counting its ${atLeast}${plural(t.subfolders, 'subfolder')}` : mail;
}

type Rule = { DisplayName?: string; Actions?: { MoveToFolder?: string; CopyToFolder?: string } };

/**
 * Rules that file mail into this folder or any folder beneath it. Losing the folder breaks them, so the
 * confirmation names them. A failed read only loses the warning, not the delete.
 */
async function rulesUsing(client: OutlookClient, ids: string[]): Promise<string[]> {
  const doomed = new Set(ids);
  try {
    const page = await client.get<{ value?: Rule[] }>(RULES_PATH);
    return (page.value ?? [])
      .filter((r) => doomed.has(r.Actions?.MoveToFolder ?? '') || doomed.has(r.Actions?.CopyToFolder ?? ''))
      .map((r) => r.DisplayName ?? '(unnamed rule)');
  } catch {
    return [];
  }
}

export function registerFolderTools(server: McpServer, client: OutlookClient): void {
  registerUngatedWrite(
    server,
    {
      description:
        'Create a mail folder: at the top level of the mailbox by default, or inside `parent`. Nothing reaches anyone else and an empty folder is easily deleted, so this does not ask for confirmation. Point rules or moves at the returned Id. ' +
        TIMEOUT_SENTENCE,
      meta: workMeta('change', { aliases: ['new folder', 'add folder', 'make folder'] }),
    },
    {
      name: 'outlook_create_folder',
      // Not gated as a single tool; in outlook_batch the batch's approval covers it.
      description: 'Create a mail folder.',
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
      inputSchema: z.object({
        name: z.string().min(1).describe('Folder name'),
        parent: z
          .string()
          .min(1)
          .optional()
          .describe('Folder id or well-known name (e.g. "inbox") to create it inside. Omit for the top level.'),
      }),
      plan: ({ name, parent }) => {
        const path = `${folderPath(parent ?? ROOT)}/childfolders`;
        const body = { DisplayName: name };
        return {
          gate: {
            action: 'mail.folder_create',
            summary: parent === undefined ? `Create folder "${name}" at the top level` : `Create folder "${name}" inside ${parent}`,
            account: undefined,
            request: { method: 'POST', path, body },
          },
          state: { name, parent, path, body },
        };
      },
      execute: async ({ name, parent, path, body }) => {
        let write;
        try {
          write = await writeOrUnknownData(() => client.write<OutlookFolder>('POST', path, body), 'outlook_list_folders');
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
        return { created: true, ...shown(write.value ?? { DisplayName: name }) };
      },
      render: minifiedResult,
      produces: {
        // An "unknown" create returned no Id: the op that needs it fails rather than guessing.
        id: (r: Record<string, unknown>) => {
          if (typeof r.Id !== 'string' || !r.Id) {
            throw new Error('outlook_create_folder did not return a folder Id (its outcome is unknown), so this op was not run.');
          }
          return r.Id;
        },
      },
    },
  );

  registerGatedWrite(server, workMeta('change', { aliases: ['rename folder', 'move folder'] }), defineWrite({
    name: 'outlook_update_folder',
    description:
      'Rename a mail folder, move it under another folder, or both. Its mail and subfolders go with it, and rules that file into it keep working. Outlook\'s own folders (Inbox, Archive, …) are refused. The result is verified by re-reading the folder. ' +
      TIMEOUT_SENTENCE,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
    inputSchema: z.object({
      id: z.string().min(1).describe('Folder Id from outlook_list_folders'),
      name: z.string().min(1).optional().describe('New name'),
      parent: z
        .string()
        .min(1)
        .optional()
        .describe('Folder id or well-known name to move it into; "msgfolderroot" is the top level'),
    }),
    plan: async ({ id, name, parent }) => {
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
      return {
        gate: {
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
          preview: { current: shown(current) },
        },
        state: { id, name, destination },
      };
    },
    execute: async ({ id, name, destination }) => {
      if (name !== undefined) {
        const rename = await writeOrUnknownData(() => client.write('PATCH', folderPath(id), { DisplayName: name }), 'outlook_list_folders');
        if (!rename.ok) return rename.result;
      }
      let finalId = id;
      if (destination !== undefined) {
        const move = await writeOrUnknownData(
          () => client.write<OutlookFolder>('POST', `${folderPath(id)}/move`, { DestinationId: destination }),
          'outlook_list_folders',
        );
        if (!move.ok) return move.result;
        finalId = move.value?.Id ?? id;
      }

      const read = await readBack(() => client.get<OutlookFolder>(folderPath(finalId)));
      if (!read.ok) return { updated: true, Id: finalId, ...unverified('outlook_list_folders') };
      const after = read.value;
      const stale = [
        ...(name !== undefined && after.DisplayName !== name ? ['name'] : []),
        ...(destination !== undefined && after.ParentFolderId !== destination ? ['location'] : []),
      ];
      return {
        updated: stale.length === 0,
        ...shown(after),
        ...(stale.length ? { warning: `Outlook accepted the write but the ${stale.join(' and ')} did not change.` } : {}),
      };
    },
    render: minifiedResult,
  }));

  server.registerTool(
    'outlook_delete_folder',
    {
      description:
        'Delete a mail folder together with every message and subfolder in it. This is permanent: the folder does NOT go to Deleted Items, so it cannot be restored from there. To keep the mail, move it out first (outlook_move_message). The confirmation totals the mail in the folder and every subfolder beneath it, and names any inbox rules that file into them (those rules stop working). Outlook\'s own folders are refused. The result is verified by re-reading the folder. A folder holding thousands of messages takes minutes to delete, so the call can come back with status "unknown" while Outlook keeps deleting: check its Total in outlook_list_folders falling, and never resend. ' +
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
      const tree = await subtree(client, { ...current, Id: current.Id ?? id });
      const rules = await rulesUsing(client, tree.ids);
      const ruleNote = rules.length
        ? ` These inbox rules file mail into it or its subfolders and will stop working: ${rules.map((r) => `"${r}"`).join(', ')}.`
        : '';
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_delete_folder',
        action: 'mail.folder_delete',
        message: `Review and confirm permanently deleting this folder with its ${contents(tree)} (it does not go to Deleted Items):${ruleNote}`,
        summary: `Permanently delete folder "${label}" with its ${contents(tree)}; it does not go to Deleted Items.${ruleNote}`,
        account: undefined,
        target: id,
        request: { method: 'DELETE', path: folderPath(id) },
        confirmToken,
        preview: { current: shown(current), ...(rules.length ? { rulesUsingIt: rules } : {}) },
      });
      if (gate) return gate;

      let write;
      try {
        write = await writeOrUnknown(() => client.write('DELETE', folderPath(id)), 'outlook_list_folders');
      } catch (e) {
        // Built-in folders such as RSS Feeds are not all reachable by a well-known name, so Outlook is the backstop.
        if (errorStatusOf(e) === 400 && /ErrorDeleteDistinguishedFolder/.test(String((e as Error).message))) {
          throw new McpToolError(`"${label}" is one of Outlook's own folders and cannot be deleted.`, {
            hint: 'Nothing was deleted. Leave it in place, or delete the folders you created inside it instead.',
          });
        }
        throw e;
      }
      if (!write.ok) return minifiedResult({ status: 'unknown', Id: id, warning: STILL_DELETING });
      const read = await readBack(() => client.get<OutlookFolder>(folderPath(id)));
      if (!read.ok && errorStatusOf(read.error) === 404) return minifiedResult({ deleted: true, Id: id });
      if (!read.ok) return minifiedResult({ deleted: true, Id: id, ...unverified('outlook_list_folders') });
      return minifiedResult({ deleted: false, Id: id, warning: 'Outlook accepted the delete but the folder is still there.' });
    },
  );
}
