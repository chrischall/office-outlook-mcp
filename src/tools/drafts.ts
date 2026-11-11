import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
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
import { composeBody, htmlAliasParam, resolveFormat, writeFormatParam } from './_body.js';

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

/**
 * How long to watch for a sent draft to leave Drafts. Outlook answers the send
 * with 202 and delivers in the background: live 2026-11-10 the draft was still
 * readable as a draft straight after the 202 although the mail arrived.
 */
const SEND_CHECK_TRIES = 8;

export interface DraftToolOptions {
  /** Pause between checks that a sent draft has left Drafts (default 1500 ms). */
  sendCheckDelayMs?: number;
}

export function registerDraftTools(server: McpServer, client: OutlookClient, opts: DraftToolOptions = {}): void {
  const delayMs = opts.sendCheckDelayMs ?? 1500;
  server.registerTool(
    'outlook_update_draft',
    {
      description:
        'Edit an unsent draft: its subject, body, and To, Cc or Bcc recipients. Only what you pass changes, and each field you pass replaces the current one: `body` replaces the WHOLE body (in a reply or forward draft that includes the quoted original), and `to: []` clears the To line. `format: "markdown"` writes the new body formatted. To add or drop a Cc without retyping the rest, use `addCc` / `removeCc` instead of `cc`: the existing Cc stays. Nothing is sent; send it with outlook_send_draft. Messages that are not drafts are refused. The confirmation shows the draft as it is now, because the replaced text cannot be recovered. The result is checked against the draft Outlook returns. ' +
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
        format: writeFormatParam.describe('How `body` is written: "text" (default), "markdown" (sent as formatted HTML) or "html". Only with `body`.'),
        html: htmlAliasParam.describe('Same as format "html" (older spelling). Only with `body`.'),
        to: recipientList.describe('New To recipients; replaces the list ([] clears it)'),
        cc: recipientList.describe('New Cc recipients; replaces the list ([] clears it)'),
        bcc: recipientList.describe('New Bcc recipients; replaces the list ([] clears it)'),
        addCc: recipientList.describe('Cc addresses to add, keeping the current Cc (not with `cc`)'),
        removeCc: recipientList.describe('Cc addresses to take off, keeping the rest (not with `cc`)'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, subject, body, format, html, to, cc, bcc, addCc, removeCc, confirmToken }, ctx) => {
      if ((html !== undefined || format !== undefined) && body === undefined) {
        throw new McpToolError(`\`${format !== undefined ? 'format' : 'html'}\` only applies together with \`body\`.`, {
          hint: 'Pass the new `body` as well.',
        });
      }
      const bodyFormat = resolveFormat(format, html);
      const editsCc = Boolean(addCc?.length || removeCc?.length);
      if (cc !== undefined && editsCc) {
        throw new McpToolError('Pass either `cc` (the whole new list) or `addCc` / `removeCc` (changes to it), not both.', {
          hint: 'Use addCc / removeCc to keep the current Cc, or cc alone to replace it.',
        });
      }
      const payload: Draft = {
        ...(subject !== undefined ? { Subject: subject } : {}),
        ...(body !== undefined ? { Body: composeBody(body, bodyFormat) } : {}),
        ...(to !== undefined ? { ToRecipients: toRecipients(to) } : {}),
        ...(cc !== undefined ? { CcRecipients: toRecipients(cc) } : {}),
        ...(bcc !== undefined ? { BccRecipients: toRecipients(bcc) } : {}),
      };
      if (Object.keys(payload).length === 0 && !editsCc) {
        throw new McpToolError('Nothing to change on this draft.', {
          hint: 'Pass at least one of subject, body, to, cc, addCc, removeCc or bcc.',
        });
      }
      const current = await readDraft(client, id);
      let ccChange = '';
      if (editsCc) {
        const now = addresses(current.CcRecipients).filter(Boolean);
        const lower = (xs: string[]) => new Set(xs.map((x) => x.toLowerCase()));
        const absent = (removeCc ?? []).filter((r) => !lower(now).has(r.toLowerCase()));
        if (absent.length) {
          throw new McpToolError(`Not on this draft's Cc line: ${absent.join(', ')}.`, {
            hint: `The current Cc is: ${now.length ? now.join(', ') : '(empty)'}.`,
          });
        }
        const dropped = lower(removeCc ?? []);
        // Entries already on the draft are kept as Outlook has them (display names included).
        const kept = (current.CcRecipients ?? []).filter((r) => !dropped.has((r.EmailAddress?.Address ?? '').toLowerCase()));
        const keptSet = lower(addresses(kept));
        const fresh: string[] = [];
        for (const a of addCc ?? []) {
          if (dropped.has(a.toLowerCase()) || keptSet.has(a.toLowerCase())) continue;
          keptSet.add(a.toLowerCase());
          fresh.push(a);
        }
        payload.CcRecipients = [...kept, ...toRecipients(fresh)];
        const next = addresses(payload.CcRecipients);
        ccChange = [
          addCc?.length ? `adding ${addCc.join(', ')}` : '',
          removeCc?.length ? `removing ${removeCc.join(', ')}` : '',
        ]
          .filter(Boolean)
          .join(' and ');
        ccChange = `Cc list (${ccChange}; it becomes: ${next.length ? next.join(', ') : '(empty)'})`;
      }
      const fields = [
        ...(subject !== undefined ? ['subject'] : []),
        ...(body !== undefined ? ['whole body'] : []),
        ...(to !== undefined ? ['To list'] : []),
        ...(cc !== undefined ? ['Cc list'] : []),
        ...(ccChange ? [ccChange] : []),
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

  server.registerTool(
    'outlook_send_draft',
    {
      description:
        'Send an existing draft as it stands: its subject, body, recipients and attachments go out unchanged. To change it first, use outlook_update_draft. Refuses a message that is not a draft, or a draft with no recipients. The confirmation shows exactly what will be sent and to whom. Once sent, Outlook files the copy in Sent Items under a new id. Outlook finishes a send in the background, so the result waits a few seconds for the draft to leave Drafts; if it has not, the result says to check Sent Items, never to resend. ' +
        TIMEOUT_SENTENCE +
        ' ' +
        CONFIRM_FLOW_SENTENCE +
        ' ' +
        CONFIRM_INJECTION_RULE,
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
      _meta: workMeta('send', { aliases: ['send draft', 'send the draft'] }),
      inputSchema: z.object({
        id: z.string().min(1).describe('Draft Id (from outlook_create_draft, outlook_reply draftOnly, or outlook_list_messages folder "drafts")'),
        confirmToken: confirmTokenParam,
      }),
    },
    async ({ id, confirmToken }, ctx) => {
      const current = await readDraft(client, id);
      const view = shown(current);
      const recipients = [...view.To, ...view.Cc, ...view.Bcc];
      if (recipients.length === 0) {
        throw new McpToolError(`Draft "${current.Subject ?? id}" has no recipients, so there is no one to send it to.`, {
          hint: 'Add To, Cc or Bcc recipients with outlook_update_draft, then send it.',
        });
      }
      const path = `${messagePath(id)}/send`;
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_send_draft',
        action: 'mail.send_draft',
        message: 'Review and confirm this email before it is sent:',
        summary: `Send draft "${current.Subject ?? ''}" to ${recipients.join(', ')}`,
        account: undefined,
        target: id,
        request: { method: 'POST', path },
        // The send carries no body: what goes out is the stored draft, so that is what is shown and bound.
        willSend: { Subject: current.Subject, To: view.To, Cc: view.Cc, Bcc: view.Bcc, Body: currentBody(current) },
        confirmToken,
      });
      if (gate) return gate;

      const write = await writeOrUnknown(() => client.write('POST', path), 'Sent Items');
      if (!write.ok) return write.result;
      const summary = { subject: current.Subject, recipients };
      // Sending moves the draft out of Drafts, so its old id stops resolving,
      // but only once Outlook has finished: watch for that for a few seconds.
      // Outlook accepted the send, so nothing here ever reports "not sent": a
      // resend on that word would mail everyone twice.
      for (let attempt = 1; attempt <= SEND_CHECK_TRIES; attempt++) {
        const read = await readBack(() => client.get<Draft>(`${messagePath(id)}?$select=Id,IsDraft`));
        if (!read.ok && errorStatusOf(read.error) === 404) return minifiedResult({ sent: true, ...summary });
        if (!read.ok) return minifiedResult({ sent: true, ...summary, ...unverified('Sent Items') });
        if (read.value.IsDraft !== true) return minifiedResult({ sent: true, ...summary });
        if (attempt < SEND_CHECK_TRIES) await new Promise((r) => setTimeout(r, delayMs));
      }
      return minifiedResult({
        sent: true,
        verified: false,
        ...summary,
        warning: `Outlook accepted the send, but the draft had not left Drafts after ${Math.round((delayMs * (SEND_CHECK_TRIES - 1)) / 1000)} s; Outlook is probably still delivering it. Check Sent Items before doing anything else, and do not send it again unless it is missing there.`,
      });
    },
  );
}
