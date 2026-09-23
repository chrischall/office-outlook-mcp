import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult } from '@chrischall/mcp-utils';
import type { OutlookClient } from '../client.js';
import { previewUnlessConfirmed, schemaConfirm } from './_confirm.js';
import { OUTBOUND_DESCRIPTION_SUFFIX } from './_untrusted.js';
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

export function registerWriteTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_send_mail',
    {
      description:
        'Send an email from the signed-in mailbox. Requires confirm:true — without it this makes no network call and returns a preview of exactly what would be sent.' +
        OUTBOUND_DESCRIPTION_SUFFIX,
      annotations: { readOnlyHint: false, destructiveHint: true },
      inputSchema: z.object({
        to: recipientList,
        cc: recipientList,
        bcc: recipientList,
        subject: z.string().describe('Subject line'),
        body: z.string().describe('Message body'),
        html: z.boolean().optional().describe('Send the body as HTML (default: plain text)'),
        saveToSentItems: z.boolean().optional().describe('Keep a copy in Sent Items (default true)'),
        confirm: schemaConfirm,
      }),
    },
    async ({ to, cc, bcc, subject, body, html, saveToSentItems, confirm }) => {
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
      const gate = previewUnlessConfirmed(
        confirm,
        `Send mail "${subject}" to ${recipients || '(no recipients)'}`,
        'POST',
        '/me/sendmail',
        payload,
      );
      if (gate) return gate;
      await client.write('POST', '/me/sendmail', payload);
      // sendmail returns 202 with an empty body; there is no id to report and
      // no resource to re-read, so say exactly what is known.
      return minifiedResult({ sent: true, subject, recipients: recipients || null });
    },
  );

  server.registerTool(
    'outlook_create_draft',
    {
      description:
        'Create a draft message in the Drafts folder without sending it. Requires confirm:true. Returns the created draft, which can be reviewed and sent from Outlook.',
      annotations: { readOnlyHint: false, destructiveHint: false },
      inputSchema: z.object({
        to: recipientList,
        cc: recipientList,
        subject: z.string().describe('Subject line'),
        body: z.string().describe('Message body'),
        html: z.boolean().optional().describe('Compose the body as HTML (default: plain text)'),
        confirm: schemaConfirm,
      }),
    },
    async ({ to, cc, subject, body, html, confirm }) => {
      const payload = {
        ...messageBody(subject, body, html),
        ToRecipients: toRecipients(to),
        ...(cc?.length ? { CcRecipients: toRecipients(cc) } : {}),
      };
      const gate = previewUnlessConfirmed(
        confirm,
        `Create draft "${subject}"`,
        'POST',
        '/me/messages',
        payload,
      );
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
        'Mark a message read or unread. Requires confirm:true. The result is verified by re-reading the message — a 2xx alone is not proof it persisted.',
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
      inputSchema: z.object({
        id: z.string().min(1).describe('Message Id'),
        isRead: z.boolean().describe('true to mark read, false to mark unread'),
        confirm: schemaConfirm,
      }),
    },
    async ({ id, isRead, confirm }) => {
      const path = `/me/messages/${encodeURIComponent(id)}`;
      const payload = { IsRead: isRead };
      const gate = previewUnlessConfirmed(
        confirm,
        `Mark message ${id} as ${isRead ? 'read' : 'unread'}`,
        'PATCH',
        path,
        payload,
      );
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
        'Move a message to another folder (e.g. "archive", "deleteditems", or a folder id from outlook_list_folders). Requires confirm:true. Moving assigns a NEW message id, which is returned.',
      annotations: { readOnlyHint: false, destructiveHint: false },
      inputSchema: z.object({
        id: z.string().min(1).describe('Message Id'),
        destination: z
          .string()
          .min(1)
          .describe('Destination folder id or well-known name (e.g. "archive")'),
        confirm: schemaConfirm,
      }),
    },
    async ({ id, destination, confirm }) => {
      const path = `/me/messages/${encodeURIComponent(id)}/move`;
      const payload = { DestinationId: destination };
      const gate = previewUnlessConfirmed(
        confirm,
        `Move message ${id} to ${destination}`,
        'POST',
        path,
        payload,
      );
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
        'Create a calendar event. Requires confirm:true. `timeZone` takes a WINDOWS zone name such as "Eastern Standard Time", not an IANA name. Attendees are emailed an invitation.' +
        OUTBOUND_DESCRIPTION_SUFFIX,
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
        attendees: recipientList,
        confirm: schemaConfirm,
      }),
    },
    async ({ subject, start, end, timeZone, location, body, attendees, confirm }) => {
      // Falls back to UTC only when the mailbox itself declares no zone.
      const tz = timeZone ?? (await mailboxTimeZone(client)) ?? 'UTC';
      const payload = {
        Subject: subject,
        Start: { DateTime: start, TimeZone: tz },
        End: { DateTime: end, TimeZone: tz },
        ...(location ? { Location: { DisplayName: location } } : {}),
        ...(body ? { Body: { ContentType: 'Text', Content: body } } : {}),
        ...(attendees?.length
          ? {
              Attendees: attendees.map((Address) => ({
                EmailAddress: { Address },
                Type: 'Required',
              })),
            }
          : {}),
      };
      const gate = previewUnlessConfirmed(
        confirm,
        `Create event "${subject}" ${start} to ${end} (${tz})`,
        'POST',
        '/me/events',
        payload,
      );
      if (gate) return gate;
      const created = await client.write<{ Id?: string; WebLink?: string }>(
        'POST',
        '/me/events',
        payload,
      );
      return minifiedResult({ created: true, Id: created?.Id, WebLink: created?.WebLink });
    },
  );
}
