import type { OutlookEvent } from '../view.js';

/**
 * The calendar item an invite message refers to, shared by the triage read
 * (outlook_get_unread) and the response write (outlook_respond_to_invite) so
 * the two cannot drift.
 */

/** The invite's calendar item, inlined on the message it arrived as. */
export const EVENT_EXPAND = 'Microsoft.OutlookServices.EventMessage/Event';

/** An invite's event: what is shown at triage and checked before responding. */
export type InviteEvent = OutlookEvent;

/**
 * The invite's event, or undefined when there is none to act on. Outlook
 * expands an invite whose calendar item is gone (declined, cancelled or
 * deleted) as an EMPTY object — live 2026-10-10: 4 of 30 meeting messages,
 * all still MeetingRequest — so an Event without an Id counts as missing.
 */
export function resolvedEvent(e: InviteEvent | null | undefined): InviteEvent | undefined {
  return e?.Id ? e : undefined;
}

/** Why an invite whose event did not resolve cannot be answered. */
export const INVITE_GONE =
  'the meeting is no longer on your calendar (already declined, cancelled, or deleted), so there is nothing to respond to';

/** Selected beside the expand to tell a meeting message from plain mail. */
export const MEETING_MESSAGE_TYPE = 'Microsoft.OutlookServices.EventMessage/MeetingMessageType';

/** True for an EventMessage (an invite, cancellation or response), by its type or `@odata.type`. */
export function isMeetingMessage(m: { MeetingMessageType?: string; '@odata.type'?: string } | undefined): boolean {
  return !!m?.MeetingMessageType || (m?.['@odata.type'] ?? '').includes('EventMessage');
}
