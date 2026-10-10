import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { McpToolError, minifiedResult, resolveView, viewParam } from '@chrischall/mcp-utils';
import { mailboxUntrusted, UNTRUSTED_DESCRIPTION_SUFFIX } from './_untrusted.js';
import type { OutlookClient } from '../client.js';
import {
  compactEvent,
  fullEvent,
  projectCollection,
  type OutlookEvent,
  VIEWS,
} from '../view.js';
import { fetchPage, nextLinkParam, plainCollection } from './_paging.js';
import { mailboxTimeZone } from '../timezone.js';

const EVENT_SELECT =
  'Id,Subject,Start,End,Location,Organizer,IsAllDay,IsCancelled,ShowAs,OnlineMeetingUrl,OnlineMeeting,BodyPreview';

/**
 * An Outlook `DateTimeTimeZone` for a caller-supplied time.
 *
 * A wall-clock time ("2026-10-12T09:00:00") is read in `zone`, matching
 * outlook_create_event. One carrying `Z` or an offset already names its
 * instant, so it is converted to UTC rather than having its offset silently
 * ignored and re-read in the mailbox zone.
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

function assertWindow(start: string, end: string): void {
  const a = Date.parse(start);
  const b = Date.parse(end);
  if (!Number.isNaN(a) && !Number.isNaN(b) && b <= a) {
    throw new McpToolError(`The window ends (${end}) before it starts (${start}).`, {
      hint: '`end` must be later than `start`.',
    });
  }
}

/** Outlook pads times to 7 fractional digits; that is noise to a reader. */
function trimTime(dt: string | undefined): string | undefined {
  return dt?.replace(/\.0+$/, '');
}

interface EmailAddressed {
  EmailAddress?: { Address?: string; Name?: string };
}

interface MeetingTimeSuggestion {
  Confidence?: number;
  OrganizerAvailability?: string;
  SuggestionReason?: string;
  MeetingTimeSlot?: {
    Start?: { DateTime?: string };
    End?: { DateTime?: string };
  };
  AttendeeAvailability?: { Availability?: string; Attendee?: EmailAddressed }[];
}

interface ScheduleInformation {
  ScheduleId?: string;
  Error?: unknown;
  ScheduleItems?: {
    Status?: string;
    Subject?: string;
    Location?: string;
    Start?: { DateTime?: string };
    End?: { DateTime?: string };
  }[];
}

const attendeeList = (what: string) =>
  z.array(z.string().min(3)).describe(`Email addresses of ${what}`);

export function registerCalendarTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_list_events',
    {
      description:
        'List calendar events in a date window. This uses the calendar VIEW, which expands recurring series into their individual occurrences — the right tool for "what is on my schedule". Times are returned in `timeZone` when given.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        start: z
          .string()
          .min(1)
          .describe('Window start, ISO 8601 (e.g. 2026-09-20T00:00:00Z)'),
        end: z.string().min(1).describe('Window end, ISO 8601'),
        timeZone: z
          .string()
          .optional()
          .describe(
            'Windows time-zone name for the returned times, e.g. "Eastern Standard Time". NOT an IANA name like America/New_York.',
          ),
        limit: z.number().int().min(1).max(200).optional().describe('Max events (default 50)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, start, end, timeZone, limit, nextLink }) => {
      const zone = timeZone ?? (await mailboxTimeZone(client));
      const data = await fetchPage<{ value?: OutlookEvent[]; '@odata.nextLink'?: string }>(
        client,
        nextLink,
        '/me/calendarview',
        {
          query: {
            startDateTime: start,
            endDateTime: end,
            $select: EVENT_SELECT,
            $orderby: 'Start/DateTime',
            $top: limit ?? 50,
          },
          ...(zone ? { prefer: `outlook.timezone="${zone}"` } : {}),
        },
      );
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return mailboxUntrusted(data);
      return mailboxUntrusted(
        projectCollection(data, v === 'full' ? fullEvent : compactEvent, 'event'),
      );
    },
  );

  server.registerTool(
    'outlook_get_event',
    {
      description:
        'Get one calendar event in full, including attendees and their responses.' +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        id: z.string().min(1).describe('Event Id from outlook_list_events'),
        timeZone: z.string().optional().describe('Windows time-zone name for returned times'),
      }),
    },
    async ({ view, id, timeZone }) => {
      const zone = timeZone ?? (await mailboxTimeZone(client));
      const data = await client.get<OutlookEvent>(
        `/me/events/${encodeURIComponent(id)}`,
        zone ? { prefer: `outlook.timezone="${zone}"` } : {},
      );
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return mailboxUntrusted(data);
      return mailboxUntrusted(v === 'compact' ? compactEvent(data) : fullEvent(data));
    },
  );

  server.registerTool(
    'outlook_list_calendars',
    {
      description: 'List the calendars in the mailbox, for picking one to read or write against.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, nextLink }) => {
      const data = await fetchPage<{
        value?: Record<string, unknown>[];
        '@odata.nextLink'?: string;
      }>(client, nextLink, '/me/calendars', { query: { $select: 'Id,Name,Color,CanEdit,Owner' } });
      if (resolveView(view, VIEWS) === 'raw') return minifiedResult(data);
      return minifiedResult(plainCollection(data));
    },
  );

  server.registerTool(
    'outlook_find_meeting_times',
    {
      description:
        "Find times when people can meet — Outlook's Scheduling Assistant. Give the attendees, a search window and a duration; Outlook checks everyone's free/busy (and the signed-in user's, as organizer) and returns ranked candidate slots with each attendee's availability. By default only working hours are searched. Use this to answer \"when can X and Y meet?\"; then book with outlook_create_event — each suggestion carries `createEventArgs` (times, zone and attendees; `subject` too when given here) to pass to it as-is. `start`/`end` are local wall-clock times in `timeZone` (default: the mailbox zone), and returned slots are in that zone too." +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        attendees: attendeeList('required attendees').min(1),
        optionalAttendees: attendeeList('optional attendees').optional(),
        subject: z
          .string()
          .optional()
          .describe("Meeting title, carried into each suggestion's createEventArgs; not sent to Outlook"),
        start: z.string().min(1).describe('Earliest the meeting may start, e.g. 2026-10-12T09:00:00'),
        end: z.string().min(1).describe('Latest the meeting may end, e.g. 2026-10-16T17:00:00'),
        durationMinutes: z.number().int().min(5).max(1440).describe('Meeting length in minutes'),
        timeZone: z
          .string()
          .optional()
          .describe('Windows time-zone name, e.g. "Eastern Standard Time". Defaults to the MAILBOX time zone.'),
        maxCandidates: z.number().int().min(1).max(50).optional().describe('Max suggestions (default 10)'),
        workingHoursOnly: z
          .boolean()
          .optional()
          .describe("Only suggest times inside working hours (default true). false searches around the clock."),
        minimumAttendeePercentage: z
          .number()
          .min(0)
          .max(100)
          .optional()
          .describe('Lowest confidence (0-100) a slot needs to be suggested. Lower it to see slots where not everyone is free.'),
        isOrganizerOptional: z
          .boolean()
          .optional()
          .describe("true when the signed-in user need not attend, so their own calendar is ignored"),
      }),
    },
    async (args) => {
      assertWindow(args.start, args.end);
      const zone = args.timeZone ?? (await mailboxTimeZone(client)) ?? 'UTC';
      const body = {
        Attendees: [
          ...args.attendees.map((Address) => ({ Type: 'Required', EmailAddress: { Address } })),
          ...(args.optionalAttendees ?? []).map((Address) => ({
            Type: 'Optional',
            EmailAddress: { Address },
          })),
        ],
        TimeConstraint: {
          ActivityDomain: args.workingHoursOnly === false ? 'Unrestricted' : 'Work',
          Timeslots: [
            { Start: dateTimeTimeZone(args.start, zone), End: dateTimeTimeZone(args.end, zone) },
          ],
        },
        MeetingDuration: `PT${args.durationMinutes}M`,
        MaxCandidates: args.maxCandidates ?? 10,
        ReturnSuggestionReasons: true,
        ...(args.minimumAttendeePercentage !== undefined
          ? { MinimumAttendeePercentage: args.minimumAttendeePercentage }
          : {}),
        ...(args.isOrganizerOptional !== undefined
          ? { IsOrganizerOptional: args.isOrganizerOptional }
          : {}),
      };
      const data = await client.post<{
        EmptySuggestionsReason?: string;
        MeetingTimeSuggestions?: MeetingTimeSuggestion[];
      }>('/me/findmeetingtimes', body, { prefer: `outlook.timezone="${zone}"` });
      if (resolveView(args.view, VIEWS) === 'raw') return mailboxUntrusted(data);
      const suggestions = (data?.MeetingTimeSuggestions ?? []).map((s) => {
        const start = trimTime(s.MeetingTimeSlot?.Start?.DateTime);
        const end = trimTime(s.MeetingTimeSlot?.End?.DateTime);
        return {
          Start: start,
          End: end,
          Confidence: s.Confidence,
          Organizer: s.OrganizerAvailability,
          Attendees: Object.fromEntries(
            (s.AttendeeAvailability ?? []).map((a) => [
              a.Attendee?.EmailAddress?.Address ?? a.Attendee?.EmailAddress?.Name ?? '?',
              a.Availability,
            ]),
          ),
          Reason: s.SuggestionReason,
          // outlook_create_event's input for this slot, so booking it is a
          // straight hand-off. The slot came back in `zone` (the Prefer header),
          // which is also the zone create reads wall-clock times in. A slot
          // without times cannot be booked, so it gets no args.
          ...(start && end
            ? {
                createEventArgs: {
                  ...(args.subject !== undefined ? { subject: args.subject } : {}),
                  start,
                  end,
                  timeZone: zone,
                  attendees: args.attendees,
                  ...(args.optionalAttendees?.length
                    ? { optionalAttendees: args.optionalAttendees }
                    : {}),
                },
              }
            : {}),
        };
      });
      return mailboxUntrusted({
        timeZone: zone,
        count: suggestions.length,
        suggestions,
        // Outlook explains an empty result ("AttendeesUnavailable", ...); an
        // empty list without it reads as "nobody is ever free".
        ...(data?.EmptySuggestionsReason
          ? { emptySuggestionsReason: data.EmptySuggestionsReason }
          : {}),
      });
    },
  );

  server.registerTool(
    'outlook_get_schedule',
    {
      description:
        "Get people's free/busy blocks for a time window — the grid view of Outlook's Scheduling Assistant. Returns each person's busy, tentative and out-of-office items (with subject and location where their calendar shares them). Use outlook_find_meeting_times to have Outlook pick slots instead. `start`/`end` are local wall-clock times in `timeZone` (default: the mailbox zone)." +
        ' ' + UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        people: attendeeList('people (or rooms) to look up').min(1).max(100),
        start: z.string().min(1).describe('Window start, e.g. 2026-10-12T09:00:00'),
        end: z.string().min(1).describe('Window end, e.g. 2026-10-12T17:00:00'),
        timeZone: z
          .string()
          .optional()
          .describe('Windows time-zone name, e.g. "Eastern Standard Time". Defaults to the MAILBOX time zone.'),
        intervalMinutes: z
          .number()
          .int()
          .min(5)
          .max(1440)
          .optional()
          .describe('Slot size for the raw view\'s AvailabilityView string (default 30)'),
      }),
    },
    async ({ view, people, start, end, timeZone, intervalMinutes }) => {
      assertWindow(start, end);
      const zone = timeZone ?? (await mailboxTimeZone(client)) ?? 'UTC';
      const data = await client.post<{ value?: ScheduleInformation[] }>(
        '/me/calendar/getschedule',
        {
          Schedules: people,
          StartTime: dateTimeTimeZone(start, zone),
          EndTime: dateTimeTimeZone(end, zone),
          AvailabilityViewInterval: intervalMinutes ?? 30,
        },
        { prefer: `outlook.timezone="${zone}"` },
      );
      if (resolveView(view, VIEWS) === 'raw') return mailboxUntrusted(data);
      const schedules = (data?.value ?? []).map((p) =>
        // An unresolvable address must not read as "free all day".
        p.Error !== undefined && p.Error !== null
          ? { Who: p.ScheduleId, Error: p.Error }
          : {
              Who: p.ScheduleId,
              Busy: (p.ScheduleItems ?? []).map((i) => ({
                Start: trimTime(i.Start?.DateTime),
                End: trimTime(i.End?.DateTime),
                Status: i.Status,
                ...(i.Subject ? { Subject: i.Subject } : {}),
                ...(i.Location ? { Location: i.Location } : {}),
              })),
            },
      );
      return mailboxUntrusted({ timeZone: zone, schedules });
    },
  );
}
