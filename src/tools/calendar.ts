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
import { mailboxTimeZone } from '../timezone.js';

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
