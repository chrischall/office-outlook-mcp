import type { OutlookEvent } from '../view.js';

/**
 * The calendar item an invite message refers to, shared by the triage read
 * (outlook_get_unread) and the response write (outlook_respond_to_invite) so
 * the two cannot drift.
 */

/** The invite's calendar item, inlined on the message it arrived as. */
export const EVENT_EXPAND = 'Microsoft.OutlookServices.EventMessage/Event';

/** An invite's event: what is shown at triage and checked before responding. */
export interface InviteEvent extends OutlookEvent {
  IsOrganizer?: boolean;
  ResponseStatus?: { Response?: string };
}
