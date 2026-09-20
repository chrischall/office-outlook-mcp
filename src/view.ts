/**
 * Response projections.
 *
 * Outlook's default projection is ~40 fields per message including the full
 * body, which is a great deal of context for an agent that is browsing. Every
 * read tool therefore takes `view` and defaults to `compact`.
 *
 * The projections below name only fields observed on live responses
 * (2026-09-20). Anything unrecognised falls through to the raw value rather
 * than being dropped — a record with holes in it is indistinguishable from
 * "there was nothing there".
 */
import { projectOrRaw } from '@chrischall/mcp-utils';

/** A `{Name, Address}` pair as Outlook nests it under `EmailAddress`. */
interface Recipient {
  EmailAddress?: { Name?: string; Address?: string };
}

function addr(r: Recipient | undefined): string | undefined {
  const e = r?.EmailAddress;
  if (!e) return undefined;
  if (e.Name && e.Address && e.Name !== e.Address) return `${e.Name} <${e.Address}>`;
  return e.Address ?? e.Name;
}

function addrs(list: Recipient[] | undefined): string[] | undefined {
  if (!Array.isArray(list) || list.length === 0) return undefined;
  return list.map((r) => addr(r)).filter((s): s is string => s !== undefined);
}

export interface OutlookMessage {
  Id?: string;
  Subject?: string;
  BodyPreview?: string;
  Body?: { ContentType?: string; Content?: string };
  From?: Recipient;
  Sender?: Recipient;
  ToRecipients?: Recipient[];
  CcRecipients?: Recipient[];
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  IsDraft?: boolean;
  HasAttachments?: boolean;
  Importance?: string;
  ConversationId?: string;
  ParentFolderId?: string;
  WebLink?: string;
  Categories?: string[];
}

export function compactMessage(m: OutlookMessage): Record<string, unknown> {
  return pruned({
    Id: m.Id,
    Subject: m.Subject,
    From: addr(m.From ?? m.Sender),
    To: addrs(m.ToRecipients),
    Received: m.ReceivedDateTime,
    IsRead: m.IsRead,
    HasAttachments: m.HasAttachments || undefined,
    Preview: m.BodyPreview?.trim() || undefined,
    // Present only when the caller asked for a single message; a listing
    // projection would otherwise carry every body.
    Body: m.Body?.Content,
  });
}

export function fullMessage(m: OutlookMessage): Record<string, unknown> {
  return pruned({
    ...compactMessage(m),
    Cc: addrs(m.CcRecipients),
    Sent: m.SentDateTime,
    Importance: m.Importance,
    IsDraft: m.IsDraft,
    ConversationId: m.ConversationId,
    ParentFolderId: m.ParentFolderId,
    Categories: m.Categories?.length ? m.Categories : undefined,
    WebLink: m.WebLink,
  });
}

export interface OutlookEvent {
  Id?: string;
  Subject?: string;
  Start?: { DateTime?: string; TimeZone?: string };
  End?: { DateTime?: string; TimeZone?: string };
  Location?: { DisplayName?: string };
  Organizer?: Recipient;
  Attendees?: (Recipient & { Status?: { Response?: string } })[];
  IsAllDay?: boolean;
  IsCancelled?: boolean;
  ShowAs?: string;
  OnlineMeetingUrl?: string;
  BodyPreview?: string;
  WebLink?: string;
}

export function compactEvent(e: OutlookEvent): Record<string, unknown> {
  return pruned({
    Id: e.Id,
    Subject: e.Subject,
    Start: e.Start?.DateTime,
    End: e.End?.DateTime,
    TimeZone: e.Start?.TimeZone,
    Location: e.Location?.DisplayName || undefined,
    Organizer: addr(e.Organizer),
    IsAllDay: e.IsAllDay || undefined,
    IsCancelled: e.IsCancelled || undefined,
  });
}

export function fullEvent(e: OutlookEvent): Record<string, unknown> {
  return pruned({
    ...compactEvent(e),
    ShowAs: e.ShowAs,
    Attendees: e.Attendees?.map((a) =>
      pruned({ Who: addr(a), Response: a.Status?.Response }),
    ),
    OnlineMeetingUrl: e.OnlineMeetingUrl,
    Preview: e.BodyPreview?.trim() || undefined,
    WebLink: e.WebLink,
  });
}

export interface OutlookFolder {
  Id?: string;
  DisplayName?: string;
  UnreadItemCount?: number;
  TotalItemCount?: number;
  ChildFolderCount?: number;
  ParentFolderId?: string;
}

export function compactFolder(f: OutlookFolder): Record<string, unknown> {
  return pruned({
    Id: f.Id,
    Name: f.DisplayName,
    Unread: f.UnreadItemCount,
    Total: f.TotalItemCount,
  });
}

/** Drop `undefined` so a compact record does not carry empty keys. */
function pruned<T extends Record<string, unknown>>(o: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(o)) {
    if (v !== undefined) out[k] = v;
  }
  return out;
}

/**
 * Apply a projection to an OData collection envelope, preserving paging.
 *
 * `projectOrRaw` hands back the raw value when the projection throws or yields
 * nothing, so a shape we did not anticipate degrades to "everything" rather
 * than to "nothing".
 */
export function projectCollection<T>(
  payload: { value?: T[]; '@odata.nextLink'?: string },
  project: (item: T) => Record<string, unknown>,
  label: string,
): Record<string, unknown> {
  const items = Array.isArray(payload?.value) ? payload.value : [];
  return pruned({
    count: items.length,
    items: items.map((item) =>
      projectOrRaw(item, project, { label, context: 'office-outlook-mcp' }),
    ),
    nextLink: payload?.['@odata.nextLink'],
  });
}
