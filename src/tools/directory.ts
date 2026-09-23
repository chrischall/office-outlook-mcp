import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, resolveView, viewParam } from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { stripOData } from '../view.js';
import { VIEWS } from './mail.js';
import { fetchPage, nextLinkParam, plainCollection } from './_paging.js';

function qs(params: Record<string, string | number | undefined>): string {
  const parts: string[] = [];
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === '') continue;
    parts.push(`${k}=${encodeURIComponent(String(v))}`);
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

type Page = { value?: Record<string, unknown>[]; '@odata.nextLink'?: string };

function collection(data: Page, raw: boolean): Record<string, unknown> {
  if (raw) return data as Record<string, unknown>;
  return plainCollection(data);
}

export function registerDirectoryTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_get_profile',
    {
      description:
        "Get the signed-in mailbox's identity (email address, display name, alias). The cheapest call that proves the credential works.",
      annotations: { readOnlyHint: true },
      inputSchema: z.object({}),
    },
    async () => minifiedResult(stripOData(await client.get<Record<string, unknown>>('/me'))),
  );

  server.registerTool(
    'outlook_get_mailbox_settings',
    {
      description:
        'Get mailbox settings: time zone, working hours, language, and automatic-replies (out-of-office) configuration.',
      annotations: { readOnlyHint: true },
      inputSchema: z.object({}),
    },
    async () =>
      minifiedResult(stripOData(await client.get<Record<string, unknown>>('/me/MailboxSettings'))),
  );

  server.registerTool(
    'outlook_list_contacts',
    {
      description: 'List saved contacts from the mailbox address book.',
      annotations: { readOnlyHint: true },
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
        `/me/contacts${qs({
          $top: limit ?? 50,
          $skip: skip,
          // `MobilePhone1`, not `MobilePhone`: the latter is the Graph name and
          // the v2.0 Contact type rejects it outright with a 400.
          $select: 'Id,DisplayName,EmailAddresses,CompanyName,JobTitle,MobilePhone1',
        })}`,
      );
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );

  server.registerTool(
    'outlook_list_people',
    {
      description:
        'List people ranked by relevance to the user — colleagues and frequent correspondents, drawn from mail traffic rather than the saved address book. Better than outlook_list_contacts for resolving "who is X" in a work mailbox.',
      annotations: { readOnlyHint: true },
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
        `/me/people${qs({
          $top: limit ?? 25,
          $select: 'Id,DisplayName,ScoredEmailAddresses,JobTitle,CompanyName',
        })}`,
      );
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );

  server.registerTool(
    'outlook_list_tasks',
    {
      description: 'List Outlook tasks with status and due dates.',
      annotations: { readOnlyHint: true },
      inputSchema: z.object({
        view: viewParam(VIEWS),
        limit: z.number().int().min(1).max(200).optional().describe('Max tasks (default 50)'),
        nextLink: nextLinkParam,
      }),
    },
    async ({ view, limit, nextLink }) => {
      const data = await fetchPage<Page>(
        client,
        nextLink,
        `/me/tasks${qs({
          $top: limit ?? 50,
          $select: 'Id,Subject,Status,Importance,DueDateTime,CompletedDateTime',
        })}`,
      );
      return minifiedResult(collection(data, resolveView(view, VIEWS) === 'raw'));
    },
  );
}
