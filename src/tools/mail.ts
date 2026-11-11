import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  expandPath,
  mapWithConcurrency,
  readEnvVar,
  type EnvSource,
  minifiedResult,
  resolveView,
  viewParam,
  McpToolError,
} from '@chrischall/mcp-utils';
import { blobStoreFromEnv, deliverFile, isHostedEnv, type BlobStore } from '@chrischall/mcp-utils/export';
import type { OutlookClient, QueryParams } from '../client.js';
import { briefDate, briefLine, briefResult, spaced, workMeta } from './_hints.js';
import {
  addr,
  addrs,
  cappedRecipients,
  compactFolder,
  compactMessage,
  fullMessage,
  projectCollection,
  type OutlookFolder,
  type OutlookMessage,
  VIEWS,
} from '../view.js';
import { EVENT_EXPAND, INVITE_GONE, MEETING_MESSAGE_TYPE, resolvedEvent, type InviteEvent } from './_invite.js';
import {
  decodeFrom,
  fetchPage,
  fromPage,
  fromParam,
  fromQuery,
  nextLinkParam,
  plainCollection,
  type FromPosition,
} from './_paging.js';
import { mailboxUntrusted, UNTRUSTED_DESCRIPTION_SUFFIX } from './_untrusted.js';
import { htmlAliasParam, readBody, readFormatParam, resolveFormat } from './_body.js';

/**
 * Well-known folder names the API accepts wherever a folder id is taken.
 * All eight verified present on a live mailbox 2026-09-20.
 */
export const WELL_KNOWN = [
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
  MEETING_MESSAGE_TYPE,
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

/** One unread message as a brief line: sender: subject — received — what it is. */
function briefUnread(m: UnreadRow): string {
  const who = m.From ?? m.Sender;
  const sender = who?.EmailAddress?.Name || who?.EmailAddress?.Address;
  const kind = triageKind(m);
  const title = sender ? `${sender}: ${m.Subject ?? '(no subject)'}` : m.Subject;
  return briefLine(title, briefDate(m.ReceivedDateTime), kind === 'mail' ? 'unread' : spaced(kind));
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
const INVITE_UNRESOLVED_HINT = `This invite's ${INVITE_GONE}.`;

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
export const ATTACHMENT_META_SELECT = 'Id,Name,ContentType,Size,IsInline';

/** Above this an attachment is never inlined, whatever its type (base64 inflates it ~4/3 again). */
const MAX_INLINE_BYTES = 5 * 1024 * 1024;

/**
 * Largest attachment whose bytes are downloaded to hand over as a link or a
 * saved file (the export tool's per-file cap too). Outlook returns them as one
 * base64 JSON string, so this bounds the memory one call can take.
 */
const MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024;

const SERVER_NAME = 'office-outlook-mcp';
const OUTPUT_DIR_ENV = 'OUTLOOK_OUTPUT_DIR';

/**
 * Where an attachment's bytes can go besides inline: the blob store (hosted —
 * a signed link), else a download directory (a local install only:
 * `OUTLOOK_OUTPUT_DIR`, else `~/Downloads/office-outlook-mcp` when
 * `~/Downloads` exists). A hosted runner's disk is never offered: nobody in
 * the caller's session can open a path on it.
 */
interface DeliveryRoutes {
  hosted: boolean;
  store?: BlobStore;
  diskDir?: string;
}

function deliveryRoutes(env: EnvSource): DeliveryRoutes {
  const store = blobStoreFromEnv(env);
  if (isHostedEnv(env)) return { hosted: true, ...(store ? { store } : {}) };
  const configured = readEnvVar(OUTPUT_DIR_ENV, { env });
  if (configured !== undefined) return { hosted: false, diskDir: expandPath(configured) };
  const downloads = join(readEnvVar('HOME', { env }) ?? homedir(), 'Downloads');
  return { hosted: false, ...(existsSync(downloads) ? { diskDir: join(downloads, SERVER_NAME) } : {}) };
}

/** Why an attachment's bytes have no way back to the caller here. */
function noRouteHint(routes: DeliveryRoutes, why: string): string {
  return routes.hosted
    ? `${why} This hosted server has no blob store configured (MCP_BLOB_BASE_URL / MCP_BLOB_SIGNING_KEY), so it cannot hand over a download link. Open it in Outlook to read it.`
    : `${why} Set ${OUTPUT_DIR_ENV} (or create ~/Downloads) to have it saved to disk, or open it in Outlook.`;
}

/** Decoded text is cut here; a longer file is marked `textTruncated`. */
const MAX_TEXT_BYTES = 200 * 1024;

/**
 * Types worth decoding to text. Senders often label a calendar or CSV file
 * `application/octet-stream`, so the file extension is consulted as well.
 */
const TEXT_TYPE = /^text\/|^application\/(json|xml|csv|ics|calendar)\b|\+(json|xml)$/i;
const TEXT_EXT = /\.(txt|csv|tsv|json|xml|ics|md|log|html?)$/i;

export interface OutlookAttachment {
  '@odata.type'?: string;
  Id?: string;
  Name?: string;
  ContentType?: string | null;
  Size?: number;
  IsInline?: boolean;
  ContentBytes?: string;
}

export type AttachmentKind = 'file' | 'item' | 'reference';

export function attachmentKind(a: OutlookAttachment): AttachmentKind {
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
export function base64DecodedLength(b64: string): number {
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

/** Read `from` once the arguments it cannot be combined with are ruled out. */
function fromPosition(from: string | undefined, conflicts: Record<string, unknown>): FromPosition | undefined {
  if (from === undefined) return undefined;
  for (const [name, value] of Object.entries(conflicts)) {
    if (value === undefined) continue;
    throw new McpToolError(`\`from\` and \`${name}\` cannot be combined.`, {
      hint:
        name === 'nextLink'
          ? 'To continue a `from` listing, pass its `next_from` as `from`; `nextLink` belongs to the newest-first listing.'
          : `\`from\` pages oldest first by time; drop \`${name}\`, or omit \`from\` for the newest-first listing.`,
    });
  }
  return decodeFrom(from);
}

/** How `from` reads, shared by both mail listings. */
const FROM_DESCRIPTION =
  'Pass `from` (an ISO 8601 UTC time, or a previous `next_from`) to page OLDEST first from that time instead: the result carries `next_from` (pass it back unchanged as `from` for the next page; on an empty page it is the same position, so it can be polled for new arrivals) and `has_more` (more already waiting), and no nextLink. Backfill: pass an old `from` and keep passing `next_from` until `has_more` is false; then keep polling with the last `next_from` for new mail.';

/** `outlook_list_messages` with `from`: one oldest-first keyset page. */
async function listFrom(
  client: OutlookClient,
  pos: FromPosition,
  opts: { view: string; folder: string | undefined; limit: number; unreadOnly: boolean | undefined },
): Promise<ReturnType<typeof mailboxUntrusted>> {
  const { bound, top } = fromQuery(pos, opts.limit);
  const data = await client.get<{ value?: OutlookMessage[]; '@odata.nextLink'?: string }>(
    `/me/mailfolders/${folderSegment(opts.folder ?? 'inbox')}/messages`,
    {
      query: {
        $top: top,
        $select: LIST_SELECT,
        // The date bound leads: Outlook rejects a $filter that does not start with the $orderby property.
        $filter: opts.unreadOnly === true ? `${bound} and IsRead eq false` : bound,
        $orderby: 'ReceivedDateTime asc',
      },
    },
  );
  const listed = Array.isArray(data?.value) ? data.value : [];
  const page = fromPage(listed, pos, opts.limit, Boolean(data?.['@odata.nextLink']));
  const { '@odata.nextLink': _superseded, ...rest } = data ?? {};
  const body =
    opts.view === 'raw'
      ? { ...rest, value: page.rows }
      : projectCollection({ value: page.rows }, opts.view === 'full' ? fullMessage : compactMessage, 'message');
  return mailboxUntrusted({ ...body, next_from: page.next_from, has_more: page.has_more });
}

export interface MailToolOptions {
  /** Environment for `OUTLOOK_OUTPUT_DIR`, `HOME` and `MCP_BLOB_*` / `MCP_DATA_DIR`. Default `process.env`. */
  env?: EnvSource;
}

export function registerMailTools(server: McpServer, client: OutlookClient, opts: MailToolOptions = {}): void {
  const env = opts.env ?? process.env;
  server.registerTool(
    'outlook_list_folders',
    {
      description:
        'List mail folders with unread and total counts. Use this to discover folder ids before listing messages; the well-known names (' +
        WELL_KNOWN.join(', ') +
        ') can be used directly without a lookup. Lists ONE level: top-level folders by default, or the child folders of `parent`. A folder with `Children` > 0 has subfolders — list them with `parent` set to its Id.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
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
        "List messages in a folder, newest first. Defaults to the inbox. Use `unreadOnly` for a triage view. Bodies are NOT included — call outlook_get_message for one. Note `search` and `unreadOnly` cannot be combined on a first-page request (the API rejects $search with $filter); pass `search` alone to search the whole mailbox. When following a `nextLink`, both are ignored, so echoing them back is harmless. Each message's To list (and Cc, in the full view) holds at most 20 addresses; a longer list adds ToCount/ToTruncated (CcCount/CcTruncated) with the full count — pass view:\"raw\" for every address. " +
        FROM_DESCRIPTION +
        ' Cannot combine with `search`, `skip` or `nextLink`.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup', { aliases: ['email', 'mail', 'inbox'] }),
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
        from: fromParam,
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, folder, limit, skip, unreadOnly, search, from, nextLink }) => {
      const pos = fromPosition(from, { nextLink, search, skip });
      if (pos !== undefined) {
        return listFrom(client, pos, { view: resolveView(view, VIEWS), folder, limit: limit ?? 25, unreadOnly });
      }
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
        'Get one message including its body. The body is requested as plain text rather than HTML (measured ~9x smaller and far easier to read); `format: "markdown"` keeps its structure (lists, tables, emphasis, links with their text) at about the same size, and `format: "html"` gives sanitised HTML. Hidden text in the HTML is dropped from both. Pass view:"raw" to get Outlook\'s untouched record. To and Cc list at most 20 addresses; a longer list adds ToCount/ToTruncated (CcCount/CcTruncated) with the full count — view:"raw" returns every address.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
      inputSchema: z.object({
        view: viewParam(VIEWS),
        id: z.string().min(1).describe('Message Id from outlook_list_messages'),
        format: readFormatParam,
        html: htmlAliasParam,
      }),
    },
    async ({ view, id, format, html }) => {
      const bodyFormat = resolveFormat(format, html);
      const data = await client.get<OutlookMessage>(`/me/messages/${encodeURIComponent(id)}`, {
        text: bodyFormat === 'text',
      });
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return mailboxUntrusted(data);
      if (data?.Body && bodyFormat !== 'text') {
        data.Body = { ...data.Body, Content: readBody(data.Body, bodyFormat) };
      }
      return mailboxUntrusted(v === 'full' ? fullMessage(data) : compactMessage(data));
    },
  );

  server.registerTool(
    'outlook_list_unread_messages',
    {
      description:
        'Triage batch: the unread messages in a folder (default inbox), newest first, each WITH its plain-text body (truncated to `maxBodyChars`) in one call — no follow-up outlook_get_message needed. Every item has a `kind`: "mail", "meetingRequest", "meetingCancelled" or "meetingResponse" (a reply to an invite you sent). Requests and cancellations also carry `event` (id, time, organizer, your current `responseStatus`) so the invite can be answered by its event id (when the meeting is no longer on your calendar — already declined, cancelled or deleted — the item has a `hint` instead and there is nothing to respond to). `to`/`cc` list at most 20 addresses; a longer list adds `toCount`/`ccCount` and `toTruncated`/`ccTruncated`. This does NOT mark anything read. Processing loop: read the batch → act on each item (reply, respond to the invite, file, flag — or leave it) → call outlook_mark_read on the plain mail you handled, so the next call returns only what is still new. An invite you answer with outlook_respond_to_invite needs no mark-read: Outlook moves it to Deleted Items under a new id, so its old id would 404. Follow `nextLink` for more. ' +
        FROM_DESCRIPTION +
        ' Cannot combine with `sinceHours` or `nextLink`.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('whats-new', { canonical: true, aliases: ['email', 'mail', 'inbox', 'unread', 'new mail', 'work email'] }),
      inputSchema: z.object({
        view: z
          .enum(['brief'])
          .optional()
          .describe(
            'Omit for the triage batch above. "brief": one line per message, `<sender>: <subject> — <received ISO time or "no date"> — <unread | meeting request | meeting cancelled | meeting response>`, at most 120 chars each, as `data` in the untrusted-content envelope; fetches no body and no invite; a line starting `-- ` is a note (`-- none` for an empty page; a title starting `-- ` or `\\` gets one leading `\\`); a last line `-- more` means another page exists (re-ask without "brief" for its nextLink / next_from).',
          ),
        folder: z.string().min(1).optional().describe('Folder id or well-known name (default "inbox")'),
        limit: z.number().int().min(1).max(50).optional().describe('Max messages (default 25, max 50)'),
        sinceHours: z
          .number()
          .positive()
          .max(24 * 365)
          .optional()
          .describe('Only mail received in the last N hours (not with `from`)'),
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
        bodyFormat: z
          .enum(['text', 'markdown'])
          .optional()
          .describe('"text" (default, Outlook\'s plain text) or "markdown" (keeps lists, tables, emphasis and links with their text; hidden text dropped)'),
        from: fromParam,
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, folder, limit, sinceHours, includeBody, maxBodyChars, bodyFormat, from, nextLink }) => {
      const asMarkdown = bodyFormat === 'markdown';
      const pos = fromPosition(from, { nextLink, sinceHours });
      const max = limit ?? 25;
      const keyset = pos !== undefined ? fromQuery(pos, max) : undefined;
      // Outlook rejects a $filter that does not lead with the $orderby
      // property ("InefficientFilter"), so the date bound goes first.
      const filter = [
        keyset?.bound,
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
            $top: keyset?.top ?? max,
            $select: UNREAD_SELECT,
            $filter: filter,
            $orderby: keyset ? 'ReceivedDateTime asc' : 'ReceivedDateTime desc',
          },
        },
      );
      const listed = Array.isArray(data?.value) ? data.value : [];
      const page = pos !== undefined ? fromPage(listed, pos, max, Boolean(data?.['@odata.nextLink'])) : undefined;
      const rows = page?.rows ?? listed;
      if (view === 'brief') {
        const more = page !== undefined ? page.has_more : Boolean(data?.['@odata.nextLink']);
        return briefResult(rows.map(briefUnread), more);
      }
      const wantBody = includeBody !== false;
      const cap = maxBodyChars ?? 4000;

      const items = await mapWithConcurrency(rows, BODY_CONCURRENCY, async (m) => {
        const kind = triageKind(m);
        const item: Record<string, unknown> = stripUndefined({
          id: m.Id,
          kind,
          receivedAt: m.ReceivedDateTime,
          from: addr(m.From ?? m.Sender),
          ...cappedRecipients('to', addrs(m.ToRecipients)),
          ...cappedRecipients('cc', addrs(m.CcRecipients)),
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
              text: wantBody && !asMarkdown,
              query: {
                $select: wantBody ? 'Id,Body' : 'Id',
                $expand: wantEvent ? EVENT_EXPAND : undefined,
              },
            },
          );
          if (wantBody) {
            const text = (readBody(one.Body, asMarkdown ? 'markdown' : 'text') ?? '').trim();
            item.body = text.length > cap ? truncateChars(text, cap) : text;
            if (text.length > cap) item.bodyTruncated = true;
          }
          if (wantEvent) {
            const event = resolvedEvent(one.Event);
            if (event) item.event = triageEvent(event);
            else item.hint = INVITE_UNRESOLVED_HINT;
          }
        } catch (e) {
          // One unreadable message must not cost the agent the whole batch.
          item.error = e instanceof Error ? e.message : String(e);
        }
        return item;
      });

      const out: Record<string, unknown> = { count: items.length, items };
      if (page !== undefined) {
        // The token supersedes the API link, which would resume newest-first.
        out.next_from = page.next_from;
        out.has_more = page.has_more;
      } else {
        const next = data?.['@odata.nextLink'];
        if (next) out.nextLink = next;
      }
      return mailboxUntrusted(out);
    },
  );

  server.registerTool(
    'outlook_list_attachments',
    {
      description:
        "List a message's attachments (name, size, content type). Metadata only — this deliberately does not download bytes.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
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
        "Get one attachment's content (ids from outlook_list_attachments). A text-like file (text/*, JSON, CSV, XML, .ics) up to 5 MB is decoded to `text`, capped at 200 KB (`textTruncated` marks a cut). An image up to 5 MB comes back as image content the model can see. Anything else (PDF, Office, archives) or larger, up to 50 MB, is handed over as the file itself: on a hosted server a signed, expiring download link (`downloadUrl`, valid until `downloadExpiresAt`) — give it to the user; on a local install a fresh file (`path`) under OUTLOOK_OUTPUT_DIR, else ~/Downloads/office-outlook-mcp. `via` says which (image, resource, link or disk). Where neither route exists the result is metadata only and `hint` says why. An attached Outlook item (a forwarded email or event) or a cloud-file link has no bytes to return; the result says which." +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
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
      const text = !image && isText(contentType, name);
      const size = meta.Size ?? 0;
      const routes = deliveryRoutes(env);
      const canHandOver = routes.store !== undefined || routes.diskDir !== undefined;
      // Decide before downloading, so a 25 MB PDF is never fetched only to be dropped.
      if (!canHandOver) {
        if (!image && !text) {
          out.hint = noRouteHint(routes, 'Binary content is not inlined: only images and text-like files are returned.');
          return mailboxUntrusted(out);
        }
        if (size > MAX_INLINE_BYTES) {
          out.hint = noRouteHint(routes, 'Content is not inlined above 5 MB.');
          return mailboxUntrusted(out);
        }
      }
      if (size > MAX_DOWNLOAD_BYTES) {
        out.hint = 'Attachments over 50 MB are not downloaded. Open it in Outlook to read it.';
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
      const decodedSize = base64DecodedLength(bytes);
      if (decodedSize > MAX_DOWNLOAD_BYTES) {
        out.hint = 'Attachments over 50 MB are not downloaded. Open it in Outlook to read it.';
        return mailboxUntrusted(out);
      }
      if (decodedSize > MAX_INLINE_BYTES && !canHandOver) {
        out.hint = noRouteHint(routes, 'Content is not inlined above 5 MB.');
        return mailboxUntrusted(out);
      }
      const buf = Buffer.from(bytes, 'base64');
      if (text && buf.length <= MAX_INLINE_BYTES) {
        out.inlined = true;
        out.text = buf.subarray(0, utf8Boundary(buf, MAX_TEXT_BYTES)).toString('utf8');
        if (buf.length > MAX_TEXT_BYTES) out.textTruncated = true;
        return mailboxUntrusted(out);
      }

      // The bytes themselves: an image inline when it fits; otherwise a signed
      // link (hosted) or a fresh 0600 file in the download directory (local).
      const delivered = await deliverFile({
        // A standalone copy: a small Buffer shares Node's allocation pool, and
        // mcp-utils' blob PUT (3.7.1) sends `view.slice().buffer` — for a Buffer
        // that is the WHOLE pool, i.e. unrelated process memory.
        bytes: new Uint8Array(buf),
        fileName: name || 'attachment',
        ...(contentType ? { mimeType: contentType } : {}),
        store: routes.store ?? null,
        keyPrefix: 'attachments',
        inlineMaxBytes: MAX_INLINE_BYTES,
        // Locally a non-image is saved rather than inlined as an embedded resource.
        prefer: image || routes.store !== undefined ? 'auto' : 'disk',
        ...(routes.diskDir !== undefined ? { disk: { dir: routes.diskDir, mode: 0o600 } } : {}),
      });
      out.inlined = delivered.via === 'image' || delivered.via === 'resource';
      out.via = delivered.via;
      if (delivered.url !== undefined) out.downloadUrl = delivered.url;
      if (delivered.expiresAt !== undefined) out.downloadExpiresAt = delivered.expiresAt;
      if (delivered.path !== undefined) out.path = delivered.path;
      if (delivered.note !== undefined) out.deliveryNote = delivered.note;
      // The name is the sender's text, so the metadata keeps its envelope; any inline bytes follow it.
      const result = mailboxUntrusted(out);
      return { ...result, content: [...result.content, ...delivered.content] };
    },
  );
}
