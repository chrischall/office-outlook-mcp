import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, resolveView, viewParam, McpToolError } from '@chrischall/mcp-utils';
import type { OutlookClient, QueryParams } from '../client.js';
import {
  compactFolder,
  compactMessage,
  fullMessage,
  projectCollection,
  type OutlookFolder,
  type OutlookMessage,
  VIEWS,
} from '../view.js';
import { fetchPage, nextLinkParam, plainCollection } from './_paging.js';
import { mailboxUntrusted, UNTRUSTED_DESCRIPTION_SUFFIX } from './_untrusted.js';

/**
 * Well-known folder names the API accepts wherever a folder id is taken.
 * All eight verified present on a live mailbox 2026-09-20.
 */
const WELL_KNOWN = [
  'inbox',
  'drafts',
  'sentitems',
  'deleteditems',
  'archive',
  'junkemail',
  'outbox',
  'clutter',
] as const;

/** Fields worth listing. Anything not here costs bytes on every row. */
const LIST_SELECT = 'Id,Subject,From,ToRecipients,ReceivedDateTime,IsRead,HasAttachments,BodyPreview';

/** `inbox` etc. pass through; anything else is treated as an opaque folder id. */
function folderSegment(folder: string): string {
  return encodeURIComponent(folder);
}

export function registerMailTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_list_folders',
    {
      description:
        'List mail folders with unread and total counts. Use this to discover folder ids before listing messages; the well-known names (' +
        WELL_KNOWN.join(', ') +
        ') can be used directly without a lookup. Lists ONE level: top-level folders by default, or the child folders of `parent`. A folder with `Children` > 0 has subfolders — list them with `parent` set to its Id.',
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        parent: z
          .string()
          .min(1)
          .optional()
          .describe('List the child folders of this folder (id or well-known name). Omit for the top level.'),
        limit: z.number().int().min(1).max(200).optional().describe('Max folders (default 50)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, parent, limit, nextLink }) => {
      const base =
        parent !== undefined ? `/me/mailfolders/${folderSegment(parent)}/childfolders` : '/me/mailfolders';
      const data = await fetchPage<{ value?: OutlookFolder[]; '@odata.nextLink'?: string }>(
        client,
        nextLink,
        base,
        { query: { $top: limit ?? 50 } },
      );
      if (resolveView(view, VIEWS) === 'raw') return minifiedResult(data);
      return minifiedResult(projectCollection(data, compactFolder, 'mail folder'));
    },
  );

  server.registerTool(
    'outlook_list_messages',
    {
      description:
        "List messages in a folder, newest first. Defaults to the inbox. Use `unreadOnly` for a triage view. Bodies are NOT included — call outlook_get_message for one. Note `search` and `unreadOnly` cannot be combined on a first-page request (the API rejects $search with $filter); pass `search` alone to search the whole mailbox. When following a `nextLink`, both are ignored, so echoing them back is harmless." +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        folder: z
          .string()
          .optional()
          .describe('Folder id or well-known name (default "inbox")'),
        limit: z.number().int().min(1).max(100).optional().describe('Max messages (default 25)'),
        skip: z.number().int().min(0).optional().describe('Offset for paging'),
        unreadOnly: z.boolean().optional().describe('Only unread messages'),
        search: z
          .string()
          .optional()
          .describe('Full-text search across the whole mailbox; cannot combine with unreadOnly on a first page (ignored with nextLink)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, folder, limit, skip, unreadOnly, search, nextLink }) => {
      // With nextLink the filters come from the link and are ignored here, so
      // the conflict only matters for a first-page request.
      if (nextLink === undefined && search !== undefined && unreadOnly === true) {
        throw new McpToolError('`search` and `unreadOnly` cannot be combined.', {
          hint: 'Outlook rejects $search together with $filter. Search first, then filter the results.',
        });
      }
      // $search always spans the mailbox; scoping it to a folder is not supported.
      const base = search !== undefined ? '/me/messages' : `/me/mailfolders/${folderSegment(folder ?? 'inbox')}/messages`;
      const params: QueryParams = {
        $top: limit ?? 25,
        $skip: skip,
        $select: LIST_SELECT,
      };
      if (search !== undefined) {
        params.$search = `"${search.replace(/"/g, '')}"`;
      } else {
        // $orderby is rejected alongside $search, so it is set only here.
        params.$orderby = 'ReceivedDateTime desc';
        if (unreadOnly === true) params.$filter = 'IsRead eq false';
      }
      const data = await fetchPage<{ value?: OutlookMessage[]; '@odata.nextLink'?: string }>(
        client,
        nextLink,
        base,
        { query: params },
      );
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return mailboxUntrusted(data);
      return mailboxUntrusted(
        projectCollection(data, v === 'full' ? fullMessage : compactMessage, 'message'),
      );
    },
  );

  server.registerTool(
    'outlook_get_message',
    {
      description:
        'Get one message including its body. The body is requested as plain text rather than HTML (measured ~9x smaller and far easier to read); pass view:"raw" to get Outlook\'s untouched record.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        id: z.string().min(1).describe('Message Id from outlook_list_messages'),
        html: z
          .boolean()
          .optional()
          .describe('Return the HTML body instead of plain text (default false)'),
      }),
    },
    async ({ view, id, html }) => {
      const data = await client.get<OutlookMessage>(`/me/messages/${encodeURIComponent(id)}`, {
        text: html !== true,
      });
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return mailboxUntrusted(data);
      return mailboxUntrusted(v === 'full' ? fullMessage(data) : compactMessage(data));
    },
  );

  server.registerTool(
    'outlook_list_attachments',
    {
      description:
        "List a message's attachments (name, size, content type). Metadata only — this deliberately does not download bytes.",
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        id: z.string().min(1).describe('Message Id'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ id, nextLink }) => {
      const data = await fetchPage<{
        value?: Record<string, unknown>[];
        '@odata.nextLink'?: string;
      }>(
        client,
        nextLink,
        `/me/messages/${encodeURIComponent(id)}/attachments`,
        { query: { $select: 'Id,Name,Size,ContentType' } },
      );
      return minifiedResult(plainCollection(data));
    },
  );
}
