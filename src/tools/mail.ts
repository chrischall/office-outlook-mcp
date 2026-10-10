import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  imageResult,
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
  type OutlookFolder,
  type OutlookMessage,
  VIEWS,
} from '../view.js';
import { EVENT_EXPAND, type InviteEvent } from './_invite.js';
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

/** Bodies (and invite events) fetched in parallel, bounded so a 50-row batch cannot burst. */
const BODY_CONCURRENCY = 4;

type TriageKind = 'mail' | 'meetingRequest' | 'meetingCancelled' | 'meetingResponse';

interface UnreadRow extends OutlookMessage {
  '@odata.type'?: string;
  MeetingMessageType?: string;
  Flag?: { FlagStatus?: string };
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

/** Set on a request/cancellation whose event did not come back with it. */
const INVITE_UNRESOLVED_HINT =
  'invite event not resolved; use outlook_get_message / outlook_respond_to_invite with messageId';

/**
 * `text` cut to at most `max` UTF-16 units, never between the halves of a
 * surrogate pair (which would leave a lone, unprintable half at the end).
 */
function truncateChars(text: string, max: number): string {
  const end = /[\uD800-\uDBFF]/.test(text.charAt(max - 1)) ? max - 1 : max;
  return text.slice(0, end);
}

function stripUndefined(o: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined));
}

/** Metadata probe for one attachment: never the bytes. */
const ATTACHMENT_META_SELECT = 'Id,Name,ContentType,Size,IsInline';

/** Above this an attachment is never inlined, whatever its type (base64 inflates it ~4/3 again). */
const MAX_INLINE_BYTES = 5 * 1024 * 1024;

/** Decoded text is cut here; a longer file is marked `textTruncated`. */
const MAX_TEXT_BYTES = 200 * 1024;

/**
 * Types worth decoding to text. Senders often label a calendar or CSV file
 * `application/octet-stream`, so the file extension is consulted as well.
 */
const TEXT_TYPE = /^text\/|^application\/(json|xml|csv|ics|calendar)\b|\+(json|xml)$/i;
const TEXT_EXT = /\.(txt|csv|tsv|json|xml|ics|md|log|html?)$/i;

interface OutlookAttachment {
  '@odata.type'?: string;
  Id?: string;
  Name?: string;
  ContentType?: string | null;
  Size?: number;
  IsInline?: boolean;
  ContentBytes?: string;
}

type AttachmentKind = 'file' | 'item' | 'reference';

function attachmentKind(a: OutlookAttachment): AttachmentKind {
  const type = a['@odata.type'] ?? '';
  if (/ItemAttachment$/.test(type)) return 'item';
  if (/ReferenceAttachment$/.test(type)) return 'reference';
  return 'file';
}

/** `image/*` renders as MCP image content; SVG is markup, so it reads as text instead. */
function isImage(contentType: string): boolean {
  return /^image\//i.test(contentType) && !/svg/i.test(contentType);
}

function isText(contentType: string, name: string): boolean {
  return TEXT_TYPE.test(contentType) || /svg/i.test(contentType) || TEXT_EXT.test(name);
}

/** Decoded size of a base64 string, without decoding it (an upper bound if it has stray characters). */
function base64DecodedLength(b64: string): number {
  const padding = b64.endsWith('==') ? 2 : b64.endsWith('=') ? 1 : 0;
  return Math.floor((b64.length * 3) / 4) - padding;
}

/**
 * The largest cut at or below `max` that does not split a UTF-8 character:
 * back off while the byte at the cut is a continuation byte (10xxxxxx).
 */
function utf8Boundary(buf: Buffer, max: number): number {
  if (buf.length <= max) return buf.length;
  let end = max;
  while (end > 0 && (buf[end] & 0xc0) === 0x80) end--;
  return end;
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
        'Triage batch: the unread messages in a folder (default inbox), newest first, each WITH its plain-text body (truncated to `maxBodyChars`) in one call — no follow-up outlook_get_message needed. Every item has a `kind`: "mail", "meetingRequest", "meetingCancelled" or "meetingResponse" (a reply to an invite you sent). Requests and cancellations also carry `event` (id, time, organizer, your current `responseStatus`) so the invite can be answered by its event id (when the event cannot be resolved the item has a `hint` instead: answer it by messageId). This does NOT mark anything read. Processing loop: read the batch → act on each item (reply, respond to the invite, file, flag — or leave it) → call outlook_mark_read on the ids you handled, so the next call returns only what is still new. Follow `nextLink` for more.' +
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
            item.body = text.length > cap ? truncateChars(text, cap) : text;
            if (text.length > cap) item.bodyTruncated = true;
          }
          if (wantEvent) {
            if (one.Event) item.event = triageEvent(one.Event);
            else item.hint = INVITE_UNRESOLVED_HINT;
          }
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

  server.registerTool(
    'outlook_get_attachment',
    {
      description:
        "Get one attachment's content (ids from outlook_list_attachments). An image comes back as image content the model can see; a text-like file (text/*, JSON, CSV, XML, .ics) is decoded to `text`, capped at 200 KB (`textTruncated` marks a cut). Any other binary (PDF, Office, archives) and anything over 5 MB is returned as metadata only — its bytes are not inlined. An attached Outlook item (a forwarded email or event) or a cloud-file link has no bytes to return; the result says which." +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        messageId: z.string().min(1).describe('Message Id'),
        attachmentId: z.string().min(1).describe('Attachment Id from outlook_list_attachments'),
      }),
    },
    async ({ messageId, attachmentId }) => {
      const path = `/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
      // Metadata first, so a 25 MB PDF is never downloaded only to be dropped.
      const meta = await client.get<OutlookAttachment>(path, { query: { $select: ATTACHMENT_META_SELECT } });
      const kind = attachmentKind(meta);
      const contentType = meta.ContentType ?? '';
      const name = meta.Name ?? '';
      const out: Record<string, unknown> = stripUndefined({
        messageId,
        attachmentId,
        name: meta.Name,
        contentType: meta.ContentType ?? undefined,
        size: meta.Size,
        isInline: meta.IsInline || undefined,
        kind,
        inlined: false,
      });

      if (kind === 'item') {
        out.hint =
          'This is an attached Outlook item (an email, event or contact), not a file, so it has no bytes to return. Ask the sender for it, or open the parent message in Outlook.';
        return mailboxUntrusted(out);
      }
      if (kind === 'reference') {
        out.hint =
          'This is a link to a cloud file (OneDrive/SharePoint), not a file stored in the message, so there are no bytes to return.';
        return mailboxUntrusted(out);
      }
      const image = isImage(contentType);
      if (!image && !isText(contentType, name)) {
        out.hint = 'Binary content is not inlined: only images and text-like files are returned. Open it in Outlook to read it.';
        return mailboxUntrusted(out);
      }
      if ((meta.Size ?? 0) > MAX_INLINE_BYTES) {
        out.hint = 'Content is not inlined above 5 MB. Open it in Outlook to read it.';
        return mailboxUntrusted(out);
      }

      const full = await client.get<OutlookAttachment>(path);
      const bytes = full.ContentBytes;
      if (typeof bytes !== 'string' || bytes.length === 0) {
        out.hint = 'Outlook returned no content for this attachment.';
        return mailboxUntrusted(out);
      }
      // Size is Outlook's figure for the whole record and can understate the
      // file; the base64 length bounds the real size, so judge it before decoding.
      if (base64DecodedLength(bytes) > MAX_INLINE_BYTES) {
        out.hint = 'Content is not inlined above 5 MB. Open it in Outlook to read it.';
        return mailboxUntrusted(out);
      }
      const buf = Buffer.from(bytes, 'base64');
      out.inlined = true;
      if (image) {
        // The name is the sender's text, so the metadata keeps its envelope; the image follows it.
        const env = mailboxUntrusted(out);
        return { ...env, content: [...env.content, ...imageResult(bytes, contentType).content] };
      }
      out.text = buf.subarray(0, utf8Boundary(buf, MAX_TEXT_BYTES)).toString('utf8');
      if (buf.length > MAX_TEXT_BYTES) out.textTruncated = true;
      return mailboxUntrusted(out);
    },
  );
}
