import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  McpToolError,
  minifiedResult,
  WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { mailboxTimeZone } from '../timezone.js';

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

/** Wall-clock arithmetic: both sides are naive times in one zone. */
const naiveMs = (dt: string) => Date.parse(`${dt.replace(/\.\d+$/, '')}Z`);
const fromNaiveMs = (ms: number) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, '');
const trimTime = (dt: string | undefined) => dt?.replace(/\.\d+$/, '');

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
      annotations: { readOnlyHint: false, destructiveHint: true },
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
      annotations: { readOnlyHint: false, destructiveHint: false },
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
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
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
    'outlook_move_message',
    {
      description:
        'Move a message to another folder (e.g. "archive", "deleteditems", or a folder id from outlook_list_folders). Moving assigns a NEW message id, which is returned.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      annotations: { readOnlyHint: false, destructiveHint: false },
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
        'Create a calendar event or meeting. A Microsoft Teams meeting is attached by default (`teamsMeeting: false` to skip) and its join link is returned. `timeZone` takes a WINDOWS zone name such as "Eastern Standard Time", not an IANA name. Attendees are emailed an invitation. To pick a time first, use outlook_find_meeting_times.' +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: false },
      inputSchema: z.object({
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
      }),
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
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
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
}
