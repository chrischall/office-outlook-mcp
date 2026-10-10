import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  mapWithConcurrency,
  minifiedResult,
  resolveView,
  viewParam,
  McpToolError,
} from '@chrischall/mcp-utils';
import type { OutlookClient, QueryParams } from '../client.js';
import {
  addr,
  addrs,
  compactFolder,
  compactMessage,
  fullMessage,
  projectCollection,
  type OutlookEvent,
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

/**
 * Fields for the triage listing. `MeetingMessageType` exists only on the
 * derived EventMessage type, so it is selected with the type cast — the same
 * form the `$expand` below uses. A plain Message row simply omits it.
 */
const UNREAD_SELECT = [
  'Id',
  'Subject',
  'From',
  'ToRecipients',
  'CcRecipients',
  'ReceivedDateTime',
  'Importance',
  'HasAttachments',
  'Categories',
  'Flag',
  'ConversationId',
  'Microsoft.OutlookServices.EventMessage/MeetingMessageType',
].join(',');

/** The invite's calendar item, inlined on the message it arrived as. */
const EVENT_EXPAND = 'Microsoft.OutlookServices.EventMessage/Event';

/** Bodies (and invite events) fetched in parallel, bounded so a 50-row batch cannot burst. */
const BODY_CONCURRENCY = 4;

type TriageKind = 'mail' | 'meetingRequest' | 'meetingCancelled' | 'meetingResponse';

interface UnreadRow extends OutlookMessage {
  '@odata.type'?: string;
  MeetingMessageType?: string;
  Flag?: { FlagStatus?: string };
}

interface InviteEvent extends OutlookEvent {
  ResponseStatus?: { Response?: string };
}

/**
 * Sort a row into what an agent does with it. `MeetingMessageType` decides;
 * an EventMessage that arrives without it (a shape we have not seen, but the
 * type name is all that is left) falls back to its `@odata.type`, so an
 * invite never silently reads as plain mail.
 */
function triageKind(m: UnreadRow): TriageKind {
  switch (m.MeetingMessageType) {
    case 'MeetingRequest':
      return 'meetingRequest';
    case 'MeetingCancelled':
      return 'meetingCancelled';
    case 'MeetingAccepted':
    case 'MeetingTenativelyAccepted': // sic — Outlook's spelling
    case 'MeetingTentativelyAccepted':
    case 'MeetingDeclined':
      return 'meetingResponse';
  }
  const type = m['@odata.type'] ?? '';
  if (!type.includes('EventMessage')) return 'mail';
  if (/Request$/.test(type)) return 'meetingRequest';
  if (/Cancel/.test(type)) return 'meetingCancelled';
  return 'meetingResponse';
}

function triageEvent(e: InviteEvent): Record<string, unknown> {
  return stripUndefined({
    id: e.Id,
    subject: e.Subject,
    start: e.Start?.DateTime,
    end: e.End?.DateTime,
    timeZone: e.Start?.TimeZone,
    location: e.Location?.DisplayName || undefined,
    organizer: addr(e.Organizer),
    responseStatus: e.ResponseStatus?.Response,
    isCancelled: e.IsCancelled,
  });
}

function stripUndefined(o: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined));
}

/** `ReceivedDateTime ge` literal: OData wants it unquoted, and seconds are precision enough. */
function sinceLiteral(hours: number): string {
  return new Date(Date.now() - hours * 3_600_000).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function registerMailTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_list_folders',
    {
      description:
        'List mail folders with unread and total counts. Use this to discover folder ids before listing messages; the well-known names (' +
        WELL_KNOWN.join(', ') +
        ') can be used directly without a lookup. Lists ONE level: top-level folders by default, or the child folders of `parent`. A folder with `Children` > 0 has subfolders — list them with `parent` set to its Id.',
      annotations: { readOnlyHint: true, openWorldHint: true },
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
      annotations: { readOnlyHint: true, openWorldHint: true },
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
      annotations: { readOnlyHint: true, openWorldHint: true },
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
    'outlook_get_unread',
    {
      description:
        'Triage batch: the unread messages in a folder (default inbox), newest first, each WITH its plain-text body (truncated to `maxBodyChars`) in one call — no follow-up outlook_get_message needed. Every item has a `kind`: "mail", "meetingRequest", "meetingCancelled" or "meetingResponse" (a reply to an invite you sent). Requests and cancellations also carry `event` (id, time, organizer, your current `responseStatus`) so the invite can be answered by its event id. This does NOT mark anything read. Processing loop: read the batch → act on each item (reply, respond to the invite, file, flag — or leave it) → call outlook_mark_read on the ids you handled, so the next call returns only what is still new. Follow `nextLink` for more.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        folder: z.string().min(1).optional().describe('Folder id or well-known name (default "inbox")'),
        limit: z.number().int().min(1).max(50).optional().describe('Max messages (default 25, max 50)'),
        sinceHours: z
          .number()
          .positive()
          .max(24 * 365)
          .optional()
          .describe('Only mail received in the last N hours'),
        includeBody: z
          .boolean()
          .optional()
          .describe('Fetch each plain-text body (default true). Invite events are resolved either way.'),
        maxBodyChars: z
          .number()
          .int()
          .min(100)
          .max(100_000)
          .optional()
          .describe('Truncate each body to this many characters (default 4000); `bodyTruncated` marks a cut'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ folder, limit, sinceHours, includeBody, maxBodyChars, nextLink }) => {
      // Outlook rejects a $filter that does not lead with the $orderby
      // property ("InefficientFilter"), so the date bound goes first.
      const filter = [
        sinceHours !== undefined ? `ReceivedDateTime ge ${sinceLiteral(sinceHours)}` : undefined,
        'IsRead eq false',
      ]
        .filter((c): c is string => c !== undefined)
        .join(' and ');
      const data = await fetchPage<{ value?: UnreadRow[]; '@odata.nextLink'?: string }>(
        client,
        nextLink,
        `/me/mailfolders/${folderSegment(folder ?? 'inbox')}/messages`,
        {
          query: {
            $top: limit ?? 25,
            $select: UNREAD_SELECT,
            $filter: filter,
            $orderby: 'ReceivedDateTime desc',
          },
        },
      );
      const rows = Array.isArray(data?.value) ? data.value : [];
      const wantBody = includeBody !== false;
      const cap = maxBodyChars ?? 4000;

      const items = await mapWithConcurrency(rows, BODY_CONCURRENCY, async (m) => {
        const kind = triageKind(m);
        const item: Record<string, unknown> = stripUndefined({
          id: m.Id,
          kind,
          receivedAt: m.ReceivedDateTime,
          from: addr(m.From ?? m.Sender),
          to: addrs(m.ToRecipients),
          cc: addrs(m.CcRecipients),
          subject: m.Subject,
          importance: m.Importance,
          hasAttachments: m.HasAttachments || undefined,
          categories: m.Categories?.length ? m.Categories : undefined,
          flag: m.Flag?.FlagStatus,
          conversationId: m.ConversationId,
        });
        // A response to an invite you sent has no event of yours to act on.
        const wantEvent = kind === 'meetingRequest' || kind === 'meetingCancelled';
        if ((!wantBody && !wantEvent) || m.Id === undefined) return item;
        try {
          const one = await client.get<UnreadRow & { Event?: InviteEvent }>(
            `/me/messages/${encodeURIComponent(m.Id)}`,
            {
              text: wantBody,
              query: {
                $select: wantBody ? 'Id,Body' : 'Id',
                $expand: wantEvent ? EVENT_EXPAND : undefined,
              },
            },
          );
          if (wantBody) {
            const text = (one.Body?.Content ?? '').trim();
            item.body = text.length > cap ? text.slice(0, cap) : text;
            if (text.length > cap) item.bodyTruncated = true;
          }
          if (wantEvent && one.Event) item.event = triageEvent(one.Event);
        } catch (e) {
          // One unreadable message must not cost the agent the whole batch.
          item.error = e instanceof Error ? e.message : String(e);
        }
        return item;
      });

      const out: Record<string, unknown> = { count: items.length, items };
      const next = data?.['@odata.nextLink'];
      if (next) out.nextLink = next;
      return mailboxUntrusted(out);
    },
  );

  server.registerTool(
    'outlook_list_attachments',
    {
      description:
        "List a message's attachments (name, size, content type). Metadata only — this deliberately does not download bytes.",
      annotations: { readOnlyHint: true, openWorldHint: true },
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
