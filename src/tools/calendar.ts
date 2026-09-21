import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, resolveView, viewParam } from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import {
  compactEvent,
  fullEvent,
  projectCollection,
  stripOData,
  type OutlookEvent,
} from '../view.js';
import { VIEWS } from './mail.js';

const EVENT_SELECT =
  'Id,Subject,Start,End,Location,Organizer,IsAllDay,IsCancelled,ShowAs,OnlineMeetingUrl,BodyPreview';

function qs(params: Record<string, string | number | undefined>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === '') continue;
    parts.push(`${k}=${encodeURIComponent(String(v))}`);
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

/**
 * The mailbox's own Windows time-zone name, or `undefined` if it cannot be read.
 *
 * Without it the calendar view answers in UTC, and a 7:15am Eastern meeting
 * comes back as 11:15 carrying nothing louder than `TimeZone: "UTC"` — read
 * past, that is a four-hour scheduling error. The mailbox already declares its
 * zone, so the default is knowable rather than guessable.
 *
 * Cached because it is a per-mailbox constant and paying a second round trip on
 * every listing to re-learn it is not worth it. Keyed by CLIENT rather than
 * held in a module variable: the zone belongs to the mailbox, not to the
 * process, so a module-global would hand one mailbox's zone to another client
 * in the same process — and would leak between tests in file order, which is
 * how a test that believes it exercised the failure path quietly stops doing so.
 *
 * A failure is swallowed deliberately — the zone is a nicety and must not take
 * the calendar down with it — and is NOT cached, so a transient blip does not
 * pin the mailbox to UTC for the life of the server.
 */
const zoneByClient = new WeakMap<OutlookClient, string>();
async function mailboxTimeZone(client: OutlookClient): Promise<string | undefined> {
  const cached = zoneByClient.get(client);
  if (cached) return cached;
  try {
    const settings = await client.get<{ TimeZone?: string }>('/me/MailboxSettings');
    const zone = settings?.TimeZone?.trim() || undefined;
    if (zone) zoneByClient.set(client, zone);
    return zone;
  } catch {
    return undefined;
  }
}

export function registerCalendarTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_list_events',
    {
      description:
        'List calendar events in a date window. This uses the calendar VIEW, which expands recurring series into their individual occurrences — the right tool for "what is on my schedule". Times are returned in `timeZone` when given.',
      annotations: { readOnlyHint: true },
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
      }),
    },
    async ({ view, start, end, timeZone, limit }) => {
      const zone = timeZone ?? (await mailboxTimeZone(client));
      const data = await client.get<{ value?: OutlookEvent[] }>(
        `/me/calendarview${qs({
          startDateTime: start,
          endDateTime: end,
          $select: EVENT_SELECT,
          $orderby: 'Start/DateTime',
          $top: limit ?? 50,
        })}`,
        zone ? { prefer: `outlook.timezone="${zone}"` } : {},
      );
      const v = resolveView(view, VIEWS);
      if (v === 'raw') return minifiedResult(data);
      return minifiedResult(
        projectCollection(data, v === 'full' ? fullEvent : compactEvent, 'event'),
      );
    },
  );

  server.registerTool(
    'outlook_get_event',
    {
      description: 'Get one calendar event in full, including attendees and their responses.',
      annotations: { readOnlyHint: true },
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
      if (v === 'raw') return minifiedResult(data);
      return minifiedResult(v === 'compact' ? compactEvent(data) : fullEvent(data));
    },
  );

  server.registerTool(
    'outlook_list_calendars',
    {
      description: 'List the calendars in the mailbox, for picking one to read or write against.',
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
      }),
    },
    async ({ view }) => {
      const data = await client.get<{ value?: Record<string, unknown>[] }>(
        `/me/calendars${qs({ $select: 'Id,Name,Color,CanEdit,Owner' })}`,
      );
      if (resolveView(view, VIEWS) === 'raw') return minifiedResult(data);
      return minifiedResult({
        count: data.value?.length ?? 0,
        items: (data.value ?? []).map(stripOData),
      });
    },
  );
}
