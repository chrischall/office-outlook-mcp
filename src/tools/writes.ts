import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  errorStatusOf,
  mapWithConcurrency,
  McpToolError,
  minifiedResult,
  WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import { isCredentialFailure, type OutlookClient } from '../client.js';
import { mailboxTimeZone } from '../timezone.js';
import { EVENT_EXPAND, type InviteEvent } from './_invite.js';
import { mailboxUntrusted, UNTRUSTED_DESCRIPTION_SUFFIX } from './_untrusted.js';

const recipientList = z
  .array(z.string().min(3))
  .optional()
  .describe('Email addresses');

function toRecipients(list: string[] | undefined): { EmailAddress: { Address: string } }[] {
  return (list ?? []).map((Address) => ({ EmailAddress: { Address } }));
}

/**
 * Body shape for a message. `ContentType` is `Text` by default — an agent
 * composing prose should not have to emit HTML, and Outlook renders text fine.
 */
function messageBody(subject: string, body: string, html: boolean | undefined) {
  return {
    Subject: subject,
    Body: { ContentType: html === true ? 'HTML' : 'Text', Content: body },
  };
}

/**
 * What makes an event a Teams meeting. Field names verified on live events
 * 2026-10-08; the join link then appears at `OnlineMeeting.JoinUrl`.
 */
const TEAMS = { IsOnlineMeeting: true, OnlineMeetingProvider: 'TeamsForBusiness' } as const;

const teamsMeetingParam = z
  .boolean()
  .optional()
  .describe('Attach a Microsoft Teams meeting (default true). false leaves it off.');

interface Attendee {
  Type?: string;
  EmailAddress?: { Address?: string; Name?: string };
}

interface StoredEvent {
  Id?: string;
  Subject?: string;
  IsOrganizer?: boolean;
  Start?: { DateTime?: string; TimeZone?: string };
  End?: { DateTime?: string; TimeZone?: string };
  Location?: { DisplayName?: string };
  Attendees?: (Attendee & { Status?: unknown })[];
  IsOnlineMeeting?: boolean;
  OnlineMeetingProvider?: string;
  OnlineMeeting?: { JoinUrl?: string } | null;
}

const EVENT_STATE_SELECT =
  'Id,Subject,IsOrganizer,Start,End,Location,Attendees,IsOnlineMeeting,OnlineMeetingProvider,OnlineMeeting';

function attendee(Address: string, Type: 'Required' | 'Optional'): Attendee {
  return { Type, EmailAddress: { Address } };
}

const lower = (a: Attendee) => a.EmailAddress?.Address?.toLowerCase();

function joinUrlOf(e: StoredEvent | undefined): string | undefined {
  return e?.OnlineMeeting?.JoinUrl || undefined;
}

function isTeams(e: StoredEvent): boolean {
  return e.IsOnlineMeeting === true && e.OnlineMeetingProvider === 'TeamsForBusiness';
}

/**
 * A `DateTimeTimeZone` for a caller-supplied time: wall-clock is read in
 * `zone`; a value with `Z` or an offset already names its instant and goes to
 * UTC rather than having the offset ignored.
 */
function dateTimeTimeZone(value: string, zone: string): { DateTime: string; TimeZone: string } {
  if (/(Z|[+-]\d{2}:?\d{2})$/i.test(value)) {
    const d = new Date(value);
    if (!Number.isNaN(d.getTime())) {
      return { DateTime: d.toISOString().replace(/\.\d{3}Z$/, ''), TimeZone: 'UTC' };
    }
  }
  return { DateTime: value, TimeZone: zone };
}

/**
 * Tool-description sentence for a write whose timeout leaves the outcome
 * unknown. Kept next to {@link unknownOutcome} so the two cannot drift.
 */
const TIMEOUT_SENTENCE =
  'If Outlook does not answer in time (or the connection drops) the result has status "unknown": the write may well have gone through, so check before retrying — never resend blindly.';

/**
 * Run a non-idempotent write; a timeout or dropped connection after the
 * request was sent (mcp-utils' WriteOutcomeUnknownError) becomes a non-error
 * "unknown" result.
 *
 * By the time the write runs the confirm token is spent, and Outlook has
 * often already accepted the request (sendmail queues with a 202). A plain
 * error invites the model to start over — fresh preview, fresh token — and
 * send a second identical email. Any other failure is a real failure and
 * still throws.
 */
async function writeOrUnknown<T>(
  run: () => Promise<T>,
  checkWhere: string,
): Promise<{ ok: true; value: T } | { ok: false; result: ReturnType<typeof minifiedResult> }> {
  try {
    return { ok: true, value: await run() };
  } catch (e) {
    if (!(e instanceof WriteOutcomeUnknownError)) throw e;
    return {
      ok: false,
      result: minifiedResult({
        status: 'unknown',
        warning: `Outlook did not confirm this write (${e.timedOut ? 'it did not answer in time' : 'the connection dropped'}), so it may already have happened. Check ${checkWhere} before retrying; do not resend blindly.`,
      }),
    };
  }
}

const INVITE_EVENT_SELECT = 'Id,Subject,IsOrganizer,IsCancelled,Start,End,Organizer,ResponseStatus';

/** Tool input → the action segment Outlook takes and the ResponseStatus it should leave. */
const INVITE_RESPONSES = {
  accept: { verb: 'accept', status: 'Accepted' },
  tentative: { verb: 'tentativelyaccept', status: 'TentativelyAccepted' },
  decline: { verb: 'decline', status: 'Declined' },
} as const;

/** Wall-clock arithmetic: both sides are naive times in one zone. */
const naiveMs = (dt: string) => Date.parse(`${dt.replace(/\.\d+$/, '')}Z`);
const fromNaiveMs = (ms: number) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, '');
const trimTime = (dt: string | undefined) => dt?.replace(/\.\d+$/, '');

/** Tool input → the action segments Outlook takes for a send and for a draft. */
const REPLY_MODES = {
  reply: { send: 'reply', draft: 'createreply', prefix: 'RE: ' },
  replyAll: { send: 'replyall', draft: 'createreplyall', prefix: 'RE: ' },
  forward: { send: 'forward', draft: 'createforward', prefix: 'FW: ' },
} as const;

type Address = { EmailAddress?: { Address?: string; Name?: string } };

interface ReplyOriginal {
  Subject?: string;
  From?: Address;
  ReplyTo?: Address[];
  ToRecipients?: Address[];
  CcRecipients?: Address[];
}

/** Addresses in order, case-insensitively de-duplicated. */
function addresses(...lists: (Address[] | undefined)[]): string[] {
  const seen = new Map<string, string>();
  for (const a of lists.flat()) {
    const v = a?.EmailAddress?.Address;
    if (v && !seen.has(v.toLowerCase())) seen.set(v.toLowerCase(), v);
  }
  return [...seen.values()];
}

const escapeHtml = (t: string) =>
  t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * The draft's body with `comment` written above the quoted original that
 * createreply/createforward put there. PATCHing a bare comment would replace
 * the quote, so the existing content is kept and the comment goes first —
 * inside `<body>` for HTML, so the result is still one document.
 */
function bodyWithComment(
  body: { ContentType?: string; Content?: string } | undefined,
  comment: string,
): { ContentType: string; Content: string } {
  const existing = body?.Content ?? '';
  if (body?.ContentType?.toLowerCase() !== 'html') {
    return { ContentType: 'Text', Content: existing ? `${comment}\n\n${existing}` : comment };
  }
  const block = `<div>${escapeHtml(comment).replace(/\r?\n/g, '<br>')}</div><br>`;
  const open = /<body[^>]*>/i.exec(existing);
  const Content = open
    ? existing.slice(0, open.index + open[0].length) + block + existing.slice(open.index + open[0].length)
    : block + existing;
  return { ContentType: 'HTML', Content };
}

/** Tool input → the `Flag.FlagStatus` Outlook stores. */
const FLAG_STATUS = { flagged: 'Flagged', complete: 'Complete', none: 'NotFlagged' } as const;

/** Messages updated in parallel, bounded so a 50-id batch cannot burst. */
const UPDATE_CONCURRENCY = 4;

interface MessageState {
  IsRead?: boolean;
  Flag?: { FlagStatus?: string };
  Categories?: string[];
}

const sameCategories = (a: string[] | undefined, b: string[] | undefined) =>
  [...(a ?? [])].map((c) => c.toLowerCase()).sort().join('\n') ===
  [...(b ?? [])].map((c) => c.toLowerCase()).sort().join('\n');

/**
 * `current` with `remove` taken out and `add` appended. Outlook matches
 * category names without regard to case, so both sides do too, and an add
 * already present (in any case) is not doubled.
 */
function editCategories(current: string[] | undefined, add: string[], remove: string[]): string[] {
  const gone = new Set(remove.map((c) => c.toLowerCase()));
  const next = (current ?? []).filter((c) => !gone.has(c.toLowerCase()));
  for (const c of add) {
    if (!next.some((n) => n.toLowerCase() === c.toLowerCase())) next.push(c);
  }
  return next;
}

/** Fields the PATCH asked for that `after` does not show. */
function unchangedFields(payload: MessageState, after: MessageState): string[] {
  const out: string[] = [];
  if (payload.IsRead !== undefined && after.IsRead !== payload.IsRead) out.push('IsRead');
  if (payload.Flag && after.Flag?.FlagStatus !== payload.Flag.FlagStatus) out.push('Flag');
  if (payload.Categories && !sameCategories(after.Categories, payload.Categories)) out.push('Categories');
  return out;
}

/**
 * outlook_create_event's input. Exported so outlook_find_meeting_times can be
 * tested to hand back `createEventArgs` this schema accepts unchanged.
 */
export const createEventInput = z.object({
  subject: z.string().describe('Event title'),
  start: z.string().min(1).describe('Start, ISO 8601 local time e.g. 2026-09-22T15:00:00'),
  end: z.string().min(1).describe('End, ISO 8601 local time'),
  timeZone: z
    .string()
    .optional()
    .describe('Windows time-zone name. Defaults to the MAILBOX time zone.'),
  location: z.string().optional().describe('Location display name'),
  body: z.string().optional().describe('Event description'),
  attendees: recipientList.describe('Required attendees\' email addresses'),
  optionalAttendees: recipientList.describe('Optional attendees\' email addresses'),
  teamsMeeting: teamsMeetingParam,
  confirmToken: confirmTokenParam,
});

export function registerWriteTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_send_mail',
    {
      description:
        'Send an email from the signed-in mailbox.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows exactly what would be sent. ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      inputSchema: z.object({
        to: recipientList,
        cc: recipientList,
        bcc: recipientList,
        subject: z.string().describe('Subject line'),
        body: z.string().describe('Message body'),
        html: z.boolean().optional().describe('Send the body as HTML (default: plain text)'),
        saveToSentItems: z.boolean().optional().describe('Keep a copy in Sent Items (default true)'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ to, cc, bcc, subject, body, html, saveToSentItems, confirmToken }, ctx) => {
      const payload = {
        Message: {
          ...messageBody(subject, body, html),
          ToRecipients: toRecipients(to),
          ...(cc?.length ? { CcRecipients: toRecipients(cc) } : {}),
          ...(bcc?.length ? { BccRecipients: toRecipients(bcc) } : {}),
        },
        SaveToSentItems: saveToSentItems !== false,
      };
      const recipients = [...(to ?? []), ...(cc ?? []), ...(bcc ?? [])].join(', ');
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_send_mail',
        action: 'mail.send',
        message: 'Review and confirm this email before it is sent:',
        summary: `Send mail "${subject}" to ${recipients || '(no recipients)'}`,
        // One signed-in mailbox per server process.
        account: undefined,
        request: { method: 'POST', path: '/me/sendmail', body: payload },
        confirmToken,
      });
      if (gate) return gate;
      const sent = await writeOrUnknown(() => client.write('POST', '/me/sendmail', payload), 'Sent Items');
      if (!sent.ok) return sent.result;
      // sendmail returns 202 with an empty body; there is no id to report and
      // no resource to re-read, so say exactly what is known.
      return minifiedResult({ sent: true, subject, recipients: recipients || null });
    },
  );

  server.registerTool(
    'outlook_create_draft',
    {
      description:
        'Create a draft message in the Drafts folder without sending it. Returns the created draft, which can be reviewed and sent from Outlook.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      inputSchema: z.object({
        to: recipientList,
        cc: recipientList,
        subject: z.string().describe('Subject line'),
        body: z.string().describe('Message body'),
        html: z.boolean().optional().describe('Compose the body as HTML (default: plain text)'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ to, cc, subject, body, html, confirmToken }, ctx) => {
      const payload = {
        ...messageBody(subject, body, html),
        ToRecipients: toRecipients(to),
        ...(cc?.length ? { CcRecipients: toRecipients(cc) } : {}),
      };
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_create_draft',
        action: 'mail.create_draft',
        message: 'Review and confirm this draft:',
        summary: `Create draft "${subject}"`,
        // One signed-in mailbox per server process.
        account: undefined,
        request: { method: 'POST', path: '/me/messages', body: payload },
        confirmToken,
      });
      if (gate) return gate;
      const created = await client.write<{ Id?: string; WebLink?: string }>(
        'POST',
        '/me/messages',
        payload,
      );
      return minifiedResult({ created: true, Id: created?.Id, WebLink: created?.WebLink });
    },
  );

  server.registerTool(
    'outlook_mark_read',
    {
      description:
        'Mark a message read or unread. The result is verified by re-reading the message — a 2xx alone is not proof it persisted.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: z.object({
        id: z.string().min(1).describe('Message Id'),
        isRead: z.boolean().describe('true to mark read, false to mark unread'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, isRead, confirmToken }, ctx) => {
      const path = `/me/messages/${encodeURIComponent(id)}`;
      const payload = { IsRead: isRead };
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_mark_read',
        action: 'mail.mark_read',
        message: 'Review and confirm this change:',
        summary: `Mark message ${id} as ${isRead ? 'read' : 'unread'}`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: id,
        request: { method: 'PATCH', path: path, body: payload },
        confirmToken,
      });
      if (gate) return gate;
      await client.write('PATCH', path, payload);
      // Re-read rather than trust the status. IsRead is the field the write
      // requested, so it is the field that proves it.
      const after = await client.get<{ IsRead?: boolean }>(`${path}?$select=IsRead`);
      return minifiedResult({
        updated: after?.IsRead === isRead,
        IsRead: after?.IsRead,
        ...(after?.IsRead === isRead ? {} : { warning: 'Outlook accepted the write but the value did not change.' }),
      });
    },
  );

  server.registerTool(
    'outlook_update_message',
    {
      description:
        'Flag, categorise and/or mark read up to 50 messages at once — the filing step of the triage loop (outlook_get_unread → act → update). `flag`: "flagged", "complete" or "none". `categories` replaces a message\'s categories; `addCategories`/`removeCategories` edit each message\'s current list instead (names from outlook_list_categories). Duplicate ids are collapsed, so each message is written once. Each id gets its own result `{ id, ok, error? }`, verified against what Outlook stored; one failure does not stop the rest, but a rejected or missing credential fails the whole call. Only your mailbox changes — nobody is notified.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: z.object({
        messageIds: z.array(z.string().min(1)).min(1).max(50).describe('Message Ids (1 to 50)'),
        flag: z.enum(['flagged', 'complete', 'none']).optional().describe('Follow-up flag to set'),
        categories: z
          .array(z.string().min(1))
          .optional()
          .describe('Replace the categories with exactly these ([] clears them)'),
        addCategories: z.array(z.string().min(1)).optional().describe('Categories to add'),
        removeCategories: z.array(z.string().min(1)).optional().describe('Categories to remove'),
        isRead: z.boolean().optional().describe('true to mark read, false to mark unread'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ messageIds, flag, categories, addCategories, removeCategories, isRead, confirmToken }, ctx) => {
      const add = addCategories ?? [];
      const remove = removeCategories ?? [];
      const editing = add.length > 0 || remove.length > 0;
      if (categories !== undefined && editing) {
        throw new McpToolError('Pass `categories` or `addCategories`/`removeCategories`, not both.', {
          hint: '`categories` replaces the whole list; the add/remove pair edits the current one.',
        });
      }
      // The part of the PATCH that is the same for every message. An add/remove
      // edit is per message, worked out from each one's current list below.
      const fixed: MessageState = {
        ...(isRead !== undefined ? { IsRead: isRead } : {}),
        ...(flag ? { Flag: { FlagStatus: FLAG_STATUS[flag] } } : {}),
        ...(categories !== undefined ? { Categories: categories } : {}),
      };
      if (Object.keys(fixed).length === 0 && !editing) {
        throw new McpToolError('Nothing to update.', {
          hint: 'Pass at least one of flag, isRead, categories, addCategories or removeCategories.',
        });
      }

      const ids = [...new Set(messageIds)];
      const pathOf = (id: string) => `/me/messages/${encodeURIComponent(id)}`;
      const changes = [
        ...(flag ? [`flag ${flag}`] : []),
        ...(isRead !== undefined ? [isRead ? 'mark read' : 'mark unread'] : []),
        ...(categories !== undefined ? [`set categories [${categories.join(', ')}]`] : []),
        ...(add.length ? [`add categories [${add.join(', ')}]`] : []),
        ...(remove.length ? [`remove categories [${remove.join(', ')}]`] : []),
      ];
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_update_message',
        action: 'mail.update',
        message: 'Review and confirm this change:',
        summary: `Update ${ids.length} message${ids.length === 1 ? '' : 's'}: ${changes.join('; ')}`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: ids.join(','),
        request: {
          method: 'PATCH',
          path: ids.length === 1 ? pathOf(ids[0]) : '/me/messages/{id}',
          body: editing
            ? { ...fixed, Categories: `(each message's current categories${add.length ? ` + ${add.join(', ')}` : ''}${remove.length ? ` - ${remove.join(', ')}` : ''})` }
            : fixed,
        },
        confirmToken,
      });
      if (gate) return gate;

      const results = await mapWithConcurrency(ids, UPDATE_CONCURRENCY, async (id) => {
        try {
          const payload: MessageState = { ...fixed };
          if (editing) {
            // PATCH replaces the whole list, so build it from the current one.
            const current = await client.get<MessageState>(`${pathOf(id)}?$select=Categories`);
            payload.Categories = editCategories(current?.Categories, add, remove);
          }
          const echoed = await client.write<MessageState | undefined>('PATCH', pathOf(id), payload);
          // Outlook answers a PATCH with the updated message; check that rather
          // than trust the 2xx, and re-read only when it came back without one.
          const after =
            echoed && typeof echoed === 'object'
              ? echoed
              : ((await client.get<MessageState>(`${pathOf(id)}?$select=IsRead,Flag,Categories`)) ?? {});
          const unchanged = unchangedFields(payload, after);
          return unchanged.length
            ? { id, ok: false, error: `Outlook accepted the write but these did not change: ${unchanged.join(', ')}.` }
            : { id, ok: true };
        } catch (e) {
          // A rejected or missing credential fails every id alike: stop and say so once.
          if (isCredentialFailure(e)) throw e;
          // PATCH is idempotent, so a failed id is safe to send again as-is.
          return { id, ok: false, error: e instanceof Error ? e.message : String(e) };
        }
      });
      const updated = results.filter((r) => r.ok).length;
      return minifiedResult({ updated, failed: results.length - updated, results });
    },
  );

  server.registerTool(
    'outlook_move_message',
    {
      description:
        'Move a message to another folder (e.g. "archive", "deleteditems", or a folder id from outlook_list_folders). Moving assigns a NEW message id, which is returned.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      inputSchema: z.object({
        id: z.string().min(1).describe('Message Id'),
        destination: z
          .string()
          .min(1)
          .describe('Destination folder id or well-known name (e.g. "archive")'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, destination, confirmToken }, ctx) => {
      const path = `/me/messages/${encodeURIComponent(id)}/move`;
      const payload = { DestinationId: destination };
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_move_message',
        action: 'mail.move',
        message: 'Review and confirm this move:',
        summary: `Move message ${id} to ${destination}`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: id,
        request: { method: 'POST', path: path, body: payload },
        confirmToken,
      });
      if (gate) return gate;
      const moved = await client.write<{ Id?: string; ParentFolderId?: string }>(
        'POST',
        path,
        payload,
      );
      return minifiedResult({
        moved: true,
        newId: moved?.Id,
        ParentFolderId: moved?.ParentFolderId,
        note: 'The message has a new Id after a move; the old one no longer resolves.',
      });
    },
  );

  server.registerTool(
    'outlook_create_event',
    {
      description:
        'Create a calendar event or meeting. A Microsoft Teams meeting is attached by default (`teamsMeeting: false` to skip) and its join link is returned. `timeZone` takes a WINDOWS zone name such as "Eastern Standard Time", not an IANA name. Attendees are emailed an invitation. To pick a time first, use outlook_find_meeting_times: each suggestion it returns carries `createEventArgs` that can be passed here as-is (add `subject` if the search did not set one).' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      inputSchema: createEventInput,
    },
    async (
      { subject, start, end, timeZone, location, body, attendees, optionalAttendees, teamsMeeting, confirmToken },
      ctx,
    ) => {
      const wantTeams = teamsMeeting !== false;
      const allAttendees = [
        ...(attendees ?? []).map((a) => attendee(a, 'Required')),
        ...(optionalAttendees ?? []).map((a) => attendee(a, 'Optional')),
      ];
      // Falls back to UTC only when the mailbox itself declares no zone; a
      // failed lookup refuses rather than guessing.
      const tz = timeZone ?? (await mailboxTimeZone(client, { required: true })) ?? 'UTC';
      const payload = {
        Subject: subject,
        Start: { DateTime: start, TimeZone: tz },
        End: { DateTime: end, TimeZone: tz },
        ...(location ? { Location: { DisplayName: location } } : {}),
        ...(body ? { Body: { ContentType: 'Text', Content: body } } : {}),
        ...(allAttendees.length ? { Attendees: allAttendees } : {}),
        ...(wantTeams ? TEAMS : {}),
      };
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_create_event',
        action: 'calendar.create_event',
        message: 'Review and confirm this event (attendees are emailed an invitation):',
        summary: `Create ${wantTeams ? 'Teams meeting' : 'event'} "${subject}" ${start} to ${end} (${tz})`,
        // One signed-in mailbox per server process.
        account: undefined,
        request: { method: 'POST', path: '/me/events', body: payload },
        confirmToken,
      });
      if (gate) return gate;
      const write = await writeOrUnknown(
        () => client.write<StoredEvent & { WebLink?: string }>('POST', '/me/events', payload),
        'the calendar (outlook_list_events)',
      );
      if (!write.ok) return write.result;
      const created = write.value;
      let joinUrl = joinUrlOf(created);
      // A 201 is not proof the Teams meeting was provisioned — the link is.
      // Re-read once in case it was filled in after the create returned.
      if (wantTeams && !joinUrl && created?.Id) {
        const after = await client.get<StoredEvent>(
          `/me/events/${encodeURIComponent(created.Id)}?$select=OnlineMeeting`,
        );
        joinUrl = joinUrlOf(after);
      }
      return minifiedResult({
        created: true,
        Id: created?.Id,
        WebLink: created?.WebLink,
        TimeZone: tz,
        ...(joinUrl ? { JoinUrl: joinUrl } : {}),
        ...(wantTeams && !joinUrl
          ? { warning: 'The event was created, but Outlook returned no Teams join link.' }
          : {}),
      });
    },
  );

  server.registerTool(
    'outlook_update_event',
    {
      description:
        'Update a meeting you organize: retitle, move, relocate, edit the description, or add/remove attendees. Moving only `start` keeps the meeting\'s length. A Microsoft Teams meeting is added if the event lacks one (`teamsMeeting: false` to skip). Attendees are sent an updated invitation. Only the organizer can update a meeting — an attendee\'s edit would change their own copy alone, so it is refused. The result is verified by re-reading the event.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
      inputSchema: z.object({
        id: z.string().min(1).describe('Event Id from outlook_list_events or outlook_get_event'),
        subject: z.string().optional().describe('New title'),
        start: z.string().optional().describe('New start, ISO 8601 local time e.g. 2026-09-22T15:00:00'),
        end: z
          .string()
          .optional()
          .describe('New end. Omit when moving `start` to keep the current length.'),
        timeZone: z
          .string()
          .optional()
          .describe('Windows time-zone name for start/end. Defaults to the MAILBOX time zone.'),
        location: z.string().optional().describe('New location display name'),
        body: z.string().optional().describe('New description (replaces the existing one)'),
        addAttendees: recipientList.describe('Required attendees to add'),
        addOptionalAttendees: recipientList.describe('Optional attendees to add'),
        removeAttendees: recipientList.describe('Attendees to remove'),
        teamsMeeting: teamsMeetingParam,
        confirmToken: confirmTokenParam,
      }),
    },
    async (args, ctx) => {
      const tz = args.timeZone ?? (await mailboxTimeZone(client, { required: true })) ?? 'UTC';
      const path = `/me/events/${encodeURIComponent(args.id)}`;
      const readEvent = (zone: string) =>
        client.get<StoredEvent>(`${path}?$select=${EVENT_STATE_SELECT}`, {
          prefer: `outlook.timezone="${zone}"`,
        });
      const current = await readEvent(tz);
      if (current?.IsOrganizer === false) {
        throw new McpToolError('Only the organizer can update this meeting.', {
          hint:
            "You are an attendee. Changing your copy would not reach the organizer or anyone else; ask the organizer, or propose a new time from Outlook.",
        });
      }

      const payload: Record<string, unknown> = {};
      if (args.subject !== undefined) payload.Subject = args.subject;
      if (args.location !== undefined) payload.Location = { DisplayName: args.location };
      if (args.body !== undefined) payload.Body = { ContentType: 'Text', Content: args.body };

      if (args.start !== undefined || args.end !== undefined) {
        const start = args.start !== undefined ? dateTimeTimeZone(args.start, tz) : undefined;
        let end = args.end !== undefined ? dateTimeTimeZone(args.end, tz) : undefined;
        // Moving the start alone keeps the length. The current times were read
        // in `tz`, so convert the duration onto whatever zone the new start uses.
        if (start && !end && current?.Start?.DateTime && current?.End?.DateTime) {
          const length = naiveMs(current.End.DateTime) - naiveMs(current.Start.DateTime);
          end = { DateTime: fromNaiveMs(naiveMs(start.DateTime) + length), TimeZone: start.TimeZone };
        }
        // A wall-clock time and one with an offset cannot be ordered without
        // the Windows zone's rules, which we do not have. Refuse rather than
        // skip the order check and write a meeting that ends before it starts.
        if (args.start !== undefined && args.end !== undefined && start?.TimeZone !== end?.TimeZone) {
          throw new McpToolError('`start` and `end` are in different zones.', {
            hint: 'Give both the same way: both local wall-clock times, or both with Z or an offset.',
          });
        }
        // Order check against the start the meeting will have. When only `end`
        // moves that is the current start, read in the end's zone.
        const startIn =
          start ??
          (end && end.TimeZone !== tz
            ? (await readEvent(end.TimeZone))?.Start
            : current?.Start);
        if (
          end &&
          startIn?.DateTime &&
          naiveMs(end.DateTime) <= naiveMs(startIn.DateTime)
        ) {
          throw new McpToolError('The meeting would end before it starts.', {
            hint: '`end` must be later than `start`.',
          });
        }
        if (start) payload.Start = start;
        if (end) payload.End = end;
      }

      const adding = [
        ...(args.addAttendees ?? []).map((a) => attendee(a, 'Required')),
        ...(args.addOptionalAttendees ?? []).map((a) => attendee(a, 'Optional')),
      ];
      const removing = new Set((args.removeAttendees ?? []).map((a) => a.toLowerCase()));
      if (adding.length || removing.size) {
        // PATCH replaces the whole list, so build it from the current one.
        const next: Attendee[] = (current?.Attendees ?? [])
          .filter((a) => !removing.has(lower(a) ?? ''))
          .map(({ Type, EmailAddress }) => ({ Type, EmailAddress }));
        for (const a of adding) {
          if (!next.some((n) => lower(n) === lower(a))) next.push(a);
        }
        payload.Attendees = next;
      }

      const addTeams = args.teamsMeeting !== false && !(current && isTeams(current));
      if (addTeams) Object.assign(payload, TEAMS);

      if (Object.keys(payload).length === 0) {
        throw new McpToolError('Nothing to update.', {
          hint: 'Pass at least one of subject, start, end, location, body, or attendee changes.',
        });
      }

      const gate = await confirmWrite(ctx, {
        tool: 'outlook_update_event',
        action: 'calendar.update_event',
        message: 'Review and confirm this change (attendees are sent an updated invitation):',
        summary: `Update "${current?.Subject ?? args.id}": ${Object.keys(payload).join(', ')}`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: args.id,
        request: { method: 'PATCH', path, body: payload },
        confirmToken: args.confirmToken,
      });
      if (gate) return gate;

      await client.write('PATCH', path, payload);

      // Re-read rather than trust the status, in the zone the write used so
      // the times compare like for like.
      // Start and End share a zone (enforced above), but either may be alone.
      const writtenZone =
        ((payload.Start ?? payload.End) as { TimeZone?: string } | undefined)?.TimeZone ?? tz;
      const after = await readEvent(writtenZone);
      const unchanged: string[] = [];
      if (payload.Subject !== undefined && after?.Subject !== payload.Subject) unchanged.push('Subject');
      for (const key of ['Start', 'End'] as const) {
        const want = payload[key] as { DateTime: string } | undefined;
        if (want && trimTime(after?.[key]?.DateTime) !== trimTime(want.DateTime)) unchanged.push(key);
      }
      if (payload.Location !== undefined && after?.Location?.DisplayName !== args.location) {
        unchanged.push('Location');
      }
      if (payload.Attendees !== undefined) {
        const want = (payload.Attendees as Attendee[]).map(lower).sort().join();
        const got = (after?.Attendees ?? []).map(lower).sort().join();
        if (want !== got) unchanged.push('Attendees');
      }
      const joinUrl = joinUrlOf(after);
      if (addTeams && !joinUrl) unchanged.push('Teams meeting');

      return minifiedResult({
        updated: unchanged.length === 0,
        Id: after?.Id ?? args.id,
        changed: Object.keys(payload),
        ...(joinUrl ? { JoinUrl: joinUrl } : {}),
        ...(unchanged.length
          ? { warning: `Outlook accepted the update but these did not change: ${unchanged.join(', ')}.` }
          : {}),
      });
    },
  );

  server.registerTool(
    'outlook_respond_to_invite',
    {
      description:
        'Accept, tentatively accept or decline a meeting invitation. Pass the invite\'s `messageId` (from outlook_get_unread, which marks it kind "meetingRequest") or the event\'s `eventId` — exactly one. The organizer is sent your response unless `sendResponse: false`, in which case only your calendar changes and no confirmation is asked. Cancelled meetings and meetings you organize are refused. The result is verified by re-reading the event\'s response status; a declined meeting may leave your calendar entirely.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
      inputSchema: z.object({
        messageId: z
          .string()
          .min(1)
          .optional()
          .describe('Id of the invite message (resolved to its event). Give this or eventId.'),
        eventId: z.string().min(1).optional().describe('Event Id. Give this or messageId.'),
        response: z.enum(['accept', 'tentative', 'decline']).describe('Your answer'),
        comment: z.string().optional().describe('Note to the organizer sent with the response'),
        sendResponse: z
          .boolean()
          .optional()
          .describe('Tell the organizer (default true). false changes only your calendar.'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ messageId, eventId, response, comment, sendResponse, confirmToken }, ctx) => {
      if ((messageId === undefined) === (eventId === undefined)) {
        throw new McpToolError('Pass exactly one of messageId or eventId.', {
          hint: 'messageId is the invite in your inbox; eventId is the meeting on your calendar.',
        });
      }

      let event: InviteEvent | undefined;
      if (messageId !== undefined) {
        // A plain message has no Event to expand, which is how a non-invite shows.
        const msg = await client.get<{ Event?: InviteEvent }>(
          `/me/messages/${encodeURIComponent(messageId)}`,
          { query: { $select: 'Id', $expand: EVENT_EXPAND } },
        );
        event = msg?.Event;
        if (!event?.Id) {
          throw new McpToolError('That message is not a meeting invite.', {
            hint: 'Only messages outlook_get_unread marks kind "meetingRequest" carry an event to answer.',
          });
        }
      } else {
        event = await client.get<InviteEvent>(
          `/me/events/${encodeURIComponent(eventId!)}?$select=${INVITE_EVENT_SELECT}`,
        );
      }
      const id = event?.Id ?? eventId!;

      if (event?.IsCancelled === true) {
        throw new McpToolError('This meeting has been cancelled; there is nothing to respond to.', {
          hint: 'Remove it from your calendar in Outlook if it is still shown.',
        });
      }
      if (event?.IsOrganizer === true || event?.ResponseStatus?.Response === 'Organizer') {
        throw new McpToolError('You organize this meeting, so you cannot respond to it.', {
          hint: 'To change or call it off, use outlook_update_event or Outlook itself.',
        });
      }

      const { verb, status } = INVITE_RESPONSES[response];
      const notify = sendResponse !== false;
      const path = `/me/events/${encodeURIComponent(id)}/${verb}`;
      const payload = { Comment: comment ?? '', SendResponse: notify };

      // Only a response the organizer receives reaches another person; without
      // one this is a change to your own calendar and needs no confirmation.
      if (notify) {
        const when = `${trimTime(event?.Start?.DateTime) ?? '?'} to ${trimTime(event?.End?.DateTime) ?? '?'}${event?.Start?.TimeZone ? ` (${event.Start.TimeZone})` : ''}`;
        const organizer = event?.Organizer?.EmailAddress?.Address ?? 'the organizer';
        const gate = await confirmWrite(ctx, {
          tool: 'outlook_respond_to_invite',
          action: 'calendar.respond_to_invite',
          message: 'Review and confirm this response (the organizer is notified):',
          summary:
            `${response === 'tentative' ? 'Tentatively accept' : response === 'accept' ? 'Accept' : 'Decline'}` +
            ` "${event?.Subject ?? id}" ${when}, organized by ${organizer}` +
            (comment ? `, with comment: "${comment}"` : ''),
          // One signed-in mailbox per server process.
          account: undefined,
          target: id,
          request: { method: 'POST', path, body: payload },
          confirmToken,
        });
        if (gate) return gate;
      }

      const write = await writeOrUnknown(
        () => client.write('POST', path, payload),
        "the event's response (outlook_get_event)",
      );
      if (!write.ok) return write.result;

      // Re-read rather than trust the 202. Declining can remove the event from
      // the calendar, so a 404 after a decline is the expected end; any other
      // failed read-back (a 5xx, a timeout) says nothing either way.
      let after: InviteEvent | undefined;
      try {
        after = await client.get<InviteEvent>(
          `/me/events/${encodeURIComponent(id)}?$select=ResponseStatus`,
        );
      } catch (e) {
        if (response === 'decline' && errorStatusOf(e) === 404) {
          return minifiedResult({
            responded: true,
            eventId: id,
            response: status,
            sentResponse: notify,
            note: 'The declined meeting is no longer on your calendar.',
          });
        }
        return minifiedResult({
          responded: null,
          eventId: id,
          sentResponse: notify,
          warning: `Outlook accepted the response but re-reading the event failed (${e instanceof Error ? e.message : String(e)}), so it is unverified.`,
        });
      }
      const got = after?.ResponseStatus?.Response;
      return minifiedResult({
        responded: got === status,
        eventId: id,
        response: got,
        sentResponse: notify,
        ...(got === status
          ? {}
          : { warning: `Outlook accepted the response but the status did not change (still ${got ?? 'unknown'}).` }),
      });
    },
  );

  server.registerTool(
    'outlook_reply',
    {
      description:
        'Reply, reply-all or forward a message, with `comment` as your new text above the quoted original. Outlook picks the recipients for reply (the sender, or its Reply-To) and reply-all (sender plus the original To and Cc, minus you); a forward needs `to`. `draftOnly: true` leaves the reply in Drafts and returns its id, sending nothing and asking no confirmation; otherwise it is sent at once.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the recipients, subject and comment. ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE +
        ' ' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true },
      inputSchema: z.object({
        messageId: z.string().min(1).describe('Id of the message to answer or forward'),
        mode: z.enum(['reply', 'replyAll', 'forward']).describe('reply (sender only), replyAll, or forward'),
        comment: z.string().describe('Your text, placed above the quoted original'),
        to: recipientList.describe('Forward recipients\' email addresses (forward only, required there)'),
        draftOnly: z
          .boolean()
          .optional()
          .describe('Save to Drafts instead of sending (default false)'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ messageId, mode, comment, to, draftOnly, confirmToken }, ctx) => {
      if (mode === 'forward' && !to?.length) {
        throw new McpToolError('A forward needs at least one recipient.', {
          hint: 'Pass `to` with the addresses to forward to.',
        });
      }
      if (mode !== 'forward' && to?.length) {
        throw new McpToolError('`to` applies only to a forward.', {
          hint: 'A reply goes where Outlook computes; use mode "forward" to send it somewhere else, or outlook_send_mail.',
        });
      }
      const { send, draft, prefix } = REPLY_MODES[mode];
      const base = `/me/messages/${encodeURIComponent(messageId)}`;
      const forwardTo = mode === 'forward' ? { ToRecipients: toRecipients(to) } : {};

      if (draftOnly === true) {
        // Nothing leaves the mailbox: the draft sits in Drafts until sent from
        // Outlook, so there is no one to protect with a confirmation.
        // The create is not idempotent: a retry after a lost answer leaves a
        // second draft, so an unconfirmed one says to look in Drafts first.
        const create = await writeOrUnknown(
          () =>
            client.write<{ Id?: string; WebLink?: string; Body?: { ContentType?: string; Content?: string } }>(
              'POST',
              `${base}/${draft}`,
              {},
            ),
          'Drafts',
        );
        if (!create.ok) return create.result;
        const created = create.value;
        if (!created?.Id) {
          throw new McpToolError('Outlook did not return the draft it created.', {
            hint: 'Check the Drafts folder before trying again.',
          });
        }
        try {
          await client.write('PATCH', `/me/messages/${encodeURIComponent(created.Id)}`, {
            Body: bodyWithComment(created.Body, comment),
            ...forwardTo,
          });
        } catch (e) {
          // The draft exists either way; an error here would hide its id and
          // invite a retry that leaves a second one.
          return minifiedResult({
            drafted: true,
            mode,
            draftId: created.Id,
            WebLink: created.WebLink,
            warning: `The draft was created but adding your comment${mode === 'forward' ? ' and recipients' : ''} failed (${e instanceof Error ? e.message : String(e)}). Edit it in Drafts, or delete it before trying again.`,
          });
        }
        return minifiedResult({ drafted: true, mode, draftId: created.Id, WebLink: created.WebLink });
      }

      // What Outlook will address the reply to, for the preview. The send
      // endpoint computes this itself; the read only shows it in advance.
      const orig = await client.get<ReplyOriginal>(base, {
        query: { $select: 'Subject,From,ReplyTo,ToRecipients,CcRecipients' },
      });
      const sender = orig?.ReplyTo?.length ? orig.ReplyTo : orig?.From ? [orig.From] : [];
      const recipients =
        mode === 'forward'
          ? { to: to ?? [], cc: [] as string[] }
          : mode === 'replyAll'
            ? { to: addresses(sender, orig?.ToRecipients), cc: addresses(orig?.CcRecipients) }
            : { to: addresses(sender), cc: [] as string[] };
      const subject = `${prefix}${orig?.Subject ?? ''}`;
      const path = `${base}/${send}`;
      const payload = { Comment: comment, ...forwardTo };
      const who = [...recipients.to, ...recipients.cc.map((c) => `cc ${c}`)].join(', ');
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_reply',
        action: `mail.${send}`,
        message: 'Review and confirm this email before it is sent:',
        summary:
          `${mode === 'forward' ? 'Forward' : mode === 'replyAll' ? 'Reply all' : 'Reply'} "${subject}"` +
          ` to ${who || '(no recipients)'}${mode === 'replyAll' ? ' (Outlook leaves out your own address)' : ''}` +
          `, with comment: "${comment}"`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: messageId,
        request: { method: 'POST', path, body: payload },
        confirmToken,
      });
      if (gate) return gate;
      const sent = await writeOrUnknown(() => client.write('POST', path, payload), 'Sent Items');
      if (!sent.ok) return sent.result;
      // Like sendmail, these return 202 with no body: nothing to re-read. The
      // subject is the original sender's text, so it goes back fenced.
      return mailboxUntrusted({ sent: true, mode, subject, to: recipients.to, cc: recipients.cc });
    },
  );
}
