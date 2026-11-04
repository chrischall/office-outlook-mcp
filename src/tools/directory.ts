import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, resolveView, viewParam } from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { briefDate, briefLine, briefResult, briefViewParam, spaced, workMeta } from './_hints.js';
import { stripOData, VIEWS } from '../view.js';
import { fetchPage, nextLinkParam, plainCollection } from './_paging.js';

type Page = { value?: Record<string, unknown>[]; '@odata.nextLink'?: string };

function collection(data: Page, raw: boolean): Record<string, unknown> {
  if (raw) return data as Record<string, unknown>;
  return plainCollection(data);
}

/** One task as a brief line: subject — due — status. */
function briefTask(t: Record<string, unknown>): string {
  const due = t.DueDateTime as { DateTime?: string } | undefined;
  const status = typeof t.Status === 'string' && t.Status ? spaced(t.Status) : 'no status';
  return briefLine(typeof t.Subject === 'string' ? t.Subject : undefined, briefDate(due?.DateTime), status);
}

export function registerDirectoryTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_get_profile',
    {
      description:
        "Get the signed-in mailbox's identity (email address, display name, alias). The cheapest call that proves the credential works.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
      inputSchema: z.object({}),
    },
    async () => minifiedResult(stripOData(await client.get<Record<string, unknown>>('/me'))),
  );

  server.registerTool(
    'outlook_get_mailbox_settings',
    {
      description:
        'Get mailbox settings: time zone, working hours, language, and automatic-replies (out-of-office) configuration.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup', { aliases: ['out of office', 'working hours'] }),
      inputSchema: z.object({}),
    },
    async () =>
      minifiedResult(stripOData(await client.get<Record<string, unknown>>('/me/MailboxSettings'))),
  );

  server.registerTool(
    'outlook_list_contacts',
    {
      description: 'List saved contacts from the mailbox address book.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup', { aliases: ['contacts', 'address book'] }),
      inputSchema: z.object({
        view: viewParam(VIEWS),
        limit: z.number().int().min(1).max(200).optional().describe('Max contacts (default 50)'),
        skip: z.number().int().min(0).optional().describe('Offset for paging'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, limit, skip, nextLink }) => {
      const data = await fetchPage<Page>(
        client,
        nextLink,
        '/me/contacts',
        {
          query: {
            $top: limit ?? 50,
            $skip: skip,
            // `MobilePhone1`, not `MobilePhone`: the latter is the Graph name and
            // the v2.0 Contact type rejects it outright with a 400.
            $select: 'Id,DisplayName,EmailAddresses,CompanyName,JobTitle,MobilePhone1',
          },
        },
      );
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );

  server.registerTool(
    'outlook_list_people',
    {
      description:
        'List people ranked by relevance to the user — colleagues and frequent correspondents, drawn from mail traffic rather than the saved address book. Better than outlook_list_contacts for resolving "who is X" in a work mailbox.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
      inputSchema: z.object({
        view: viewParam(VIEWS),
        limit: z.number().int().min(1).max(100).optional().describe('Max people (default 25)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, limit, nextLink }) => {
      const data = await fetchPage<Page>(
        client,
        nextLink,
        '/me/people',
        {
          query: {
            $top: limit ?? 25,
            $select: 'Id,DisplayName,ScoredEmailAddresses,JobTitle,CompanyName',
          },
        },
      );
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );

  server.registerTool(
    'outlook_list_tasks',
    {
      description: 'List Outlook tasks with status and due dates.',
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('whats-due', { canonical: true, aliases: ['tasks', 'to do', 'to-dos', 'due', 'deadlines'] }),
      inputSchema: z.object({
        view: briefViewParam(VIEWS, 'A brief status is the task\'s own (not started, in progress, completed, …).'),
        limit: z.number().int().min(1).max(200).optional().describe('Max tasks (default 50)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, limit, nextLink }) => {
      const data = await fetchPage<Page>(
        client,
        nextLink,
        '/me/tasks',
        {
          query: {
            $top: limit ?? 50,
            $select: 'Id,Subject,Status,Importance,DueDateTime,CompletedDateTime',
          },
        },
      );
      if (view === 'brief') return briefResult((data?.value ?? []).map(briefTask), Boolean(data?.['@odata.nextLink']));
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );

  server.registerTool(
    'outlook_list_categories',
    {
      description:
        "List the mailbox's categories (the master list) with their colours — the names outlook_update_message can add or remove.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      _meta: workMeta('lookup'),
      inputSchema: z.object({}),
    },
    async () => {
      const data = await client.get<{ value?: { DisplayName?: string; Color?: string }[] }>(
        '/me/outlook/masterCategories',
      );
      const categories = (data?.value ?? []).map((c) => ({ name: c.DisplayName, color: c.Color }));
      return minifiedResult({ count: categories.length, categories });
    },
  );
}
