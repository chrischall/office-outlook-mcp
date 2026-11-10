import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  confirmTokenParam,
  confirmWrite,
  errorStatusOf,
  McpToolError,
  minifiedResult,
} from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { workMeta } from './_hints.js';
import { readBack, unverified } from './rules.js';
import { recipientList, TIMEOUT_SENTENCE, toRecipients, writeOrUnknown } from './writes.js';

type Recipient = { EmailAddress?: { Address?: string } };
type Draft = {
  Id?: string;
  IsDraft?: boolean;
  Subject?: string;
  Body?: { ContentType?: string; Content?: string };
  ToRecipients?: Recipient[];
  CcRecipients?: Recipient[];
  BccRecipients?: Recipient[];
  WebLink?: string;
};

const DRAFT_SELECT = 'Id,IsDraft,Subject,Body,ToRecipients,CcRecipients,BccRecipients,WebLink';
const messagePath = (id: string) => `/me/messages/${encodeURIComponent(id)}`;

/** The preview shows the body being replaced; past this it is cut, and says so. */
const PREVIEW_BODY_MAX = 2000;

const addresses = (list: Recipient[] | undefined) => (list ?? []).map((r) => r.EmailAddress?.Address ?? '');
const sameAddresses = (a: Recipient[] | undefined, b: Recipient[] | undefined) =>
  addresses(a).map((x) => x.toLowerCase()).sort().join('\n') === addresses(b).map((x) => x.toLowerCase()).sort().join('\n');

/** Outlook may store line breaks as CRLF; the words are what must match. */
const sameText = (a: string | undefined, b: string | undefined) =>
  (a ?? '').replace(/\r\n/g, '\n').trim() === (b ?? '').replace(/\r\n/g, '\n').trim();

function shown(d: Draft) {
  return {
    Id: d.Id,
    Subject: d.Subject,
    To: addresses(d.ToRecipients),
    Cc: addresses(d.CcRecipients),
    Bcc: addresses(d.BccRecipients),
    WebLink: d.WebLink,
  };
}

function currentBody(d: Draft): string {
  const content = d.Body?.Content ?? '';
  return content.length > PREVIEW_BODY_MAX ? `${content.slice(0, PREVIEW_BODY_MAX)}… (${content.length} characters in all)` : content;
}

/** Fields the PATCH set that the draft Outlook returned does not show. */
function unchanged(payload: Draft, after: Draft): string[] {
  const out: string[] = [];
  if (payload.Subject !== undefined && after.Subject !== payload.Subject) out.push('Subject');
  if (payload.Body) {
    // Outlook wraps an HTML body in its own <html><head>…, so only its type can be checked.
    const ok =
      payload.Body.ContentType === 'HTML'
        ? after.Body?.ContentType === 'HTML'
        : after.Body?.ContentType === 'Text' && sameText(after.Body?.Content, payload.Body.Content);
    if (!ok) out.push('Body');
  }
  if (payload.ToRecipients && !sameAddresses(after.ToRecipients, payload.ToRecipients)) out.push('To');
  if (payload.CcRecipients && !sameAddresses(after.CcRecipients, payload.CcRecipients)) out.push('Cc');
  if (payload.BccRecipients && !sameAddresses(after.BccRecipients, payload.BccRecipients)) out.push('Bcc');
  return out;
}

async function readDraft(client: OutlookClient, id: string): Promise<Draft> {
  let found: Draft;
  try {
    found = await client.get<Draft>(`${messagePath(id)}?$select=${DRAFT_SELECT}`);
  } catch (e) {
    const status = errorStatusOf(e);
    if (status !== 400 && status !== 404) throw e;
    throw new McpToolError(`No message matches "${id}".`, {
      hint: 'Pass a draft Id from outlook_create_draft, outlook_reply with draftOnly, or outlook_list_messages with folder "drafts".',
    });
  }
  if (found.IsDraft !== true) {
    throw new McpToolError(`"${found.Subject ?? id}" is not a draft, so it cannot be edited.`, {
      hint: 'Only unsent drafts can be changed. To send something different, create a new draft or use outlook_send_message.',
    });
  }
  return found;
}

export function registerDraftTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_update_draft',
    {
      description:
        'Edit an unsent draft: its subject, body, and To, Cc or Bcc recipients. Only what you pass changes, and each field you pass replaces the current one: `body` replaces the WHOLE body (in a reply or forward draft that includes the quoted original), and `to: []` clears the To line. Nothing is sent; send the draft from Outlook. Messages that are not drafts are refused. The confirmation shows the draft as it is now, because the replaced text cannot be recovered. The result is checked against the draft Outlook returns. ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_FLOW_SENTENCE,
      // Overwrites text that is not kept anywhere else, so it is destructive
      // even though it reaches nobody. Repeating the same edit changes nothing.
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
      _meta: workMeta('change', { aliases: ['edit draft', 'change draft', 'rewrite draft'] }),
      inputSchema: z.object({
        id: z.string().min(1).describe('Draft Id (from outlook_create_draft, outlook_reply draftOnly, or outlook_list_messages folder "drafts")'),
        subject: z.string().optional().describe('New subject line'),
        body: z.string().optional().describe('New body; replaces the whole current body'),
        html: z.boolean().optional().describe('Treat `body` as HTML (default: plain text). Only with `body`.'),
        to: recipientList.describe('New To recipients; replaces the list ([] clears it)'),
        cc: recipientList.describe('New Cc recipients; replaces the list ([] clears it)'),
        bcc: recipientList.describe('New Bcc recipients; replaces the list ([] clears it)'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, subject, body, html, to, cc, bcc, confirmToken }, ctx) => {
      if (html !== undefined && body === undefined) {
        throw new McpToolError('`html` only applies together with `body`.', { hint: 'Pass the new `body` as well.' });
      }
      const payload: Draft = {
        ...(subject !== undefined ? { Subject: subject } : {}),
        ...(body !== undefined ? { Body: { ContentType: html === true ? 'HTML' : 'Text', Content: body } } : {}),
        ...(to !== undefined ? { ToRecipients: toRecipients(to) } : {}),
        ...(cc !== undefined ? { CcRecipients: toRecipients(cc) } : {}),
        ...(bcc !== undefined ? { BccRecipients: toRecipients(bcc) } : {}),
      };
      if (Object.keys(payload).length === 0) {
        throw new McpToolError('Nothing to change on this draft.', {
          hint: 'Pass at least one of subject, body, to, cc or bcc.',
        });
      }
      const current = await readDraft(client, id);
      const fields = [
        ...(subject !== undefined ? ['subject'] : []),
        ...(body !== undefined ? ['whole body'] : []),
        ...(to !== undefined ? ['To list'] : []),
        ...(cc !== undefined ? ['Cc list'] : []),
        ...(bcc !== undefined ? ['Bcc list'] : []),
      ];
      const path = messagePath(id);
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_update_draft',
        action: 'mail.draft_update',
        message: `Review and confirm this draft edit. It will replace the draft's ${fields.join(', ')}; the current draft is shown below. Nothing is sent.`,
        summary: `Edit draft "${current.Subject ?? id}": replace its ${fields.join(', ')}. Nothing is sent.`,
        account: undefined,
        target: id,
        request: { method: 'PATCH', path, body: payload },
        confirmToken,
        preview: { current: { ...shown(current), Body: currentBody(current) } },
      });
      if (gate) return gate;

      const write = await writeOrUnknown(() => client.write<Draft | undefined>('PATCH', path, payload), 'Drafts');
      if (!write.ok) return write.result;
      // Outlook answers the PATCH with the updated draft; re-read only when it did not.
      const echoed = write.value;
      let after: Draft;
      if (echoed && typeof echoed === 'object' && echoed.Id) {
        after = echoed;
      } else {
        const read = await readBack(() => client.get<Draft>(`${path}?$select=${DRAFT_SELECT}`));
        if (!read.ok) return minifiedResult({ updated: true, Id: id, ...unverified('the draft in Drafts') });
        after = read.value;
      }
      const stale = unchanged(payload, after);
      return minifiedResult({
        updated: stale.length === 0,
        ...shown(after),
        ...(stale.length ? { warning: `Outlook accepted the edit but these did not change: ${stale.join(', ')}.` } : {}),
      });
    },
  );
}
