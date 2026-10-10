import { describe, expect, it } from 'vitest';
import {
  compactEvent,
  compactFolder,
  compactMessage,
  fullEvent,
  fullMessage,
  MAX_RECIPIENTS,
  projectCollection,
  VIEWS,
} from '../src/view.js';

/** A message shaped like the live 2026-09-20 capture. */
const message = {
  Id: 'm1',
  Subject: 'Quarterly review',
  BodyPreview: '  Hello there  ',
  Body: { ContentType: 'Text', Content: 'Hello there, long body' },
  From: { EmailAddress: { Name: 'Alice Smith', Address: 'alice@example.com' } },
  ToRecipients: [{ EmailAddress: { Name: 'Bob', Address: 'bob@example.com' } }],
  CcRecipients: [{ EmailAddress: { Address: 'carol@example.com' } }],
  ReceivedDateTime: '2026-09-20T10:00:00Z',
  SentDateTime: '2026-09-20T09:59:00Z',
  IsRead: false,
  HasAttachments: false,
  ConversationId: 'c1',
  WebLink: 'https://outlook.office.com/x',
};

describe('message projection', () => {
  it('renders a recipient as "Name <address>" and trims the preview', () => {
    const c = compactMessage(message);
    expect(c.From).toBe('Alice Smith <alice@example.com>');
    expect(c.To).toEqual(['Bob <bob@example.com>']);
    expect(c.Preview).toBe('Hello there');
  });

  it('falls back to the address alone when there is no distinct name', () => {
    expect(
      compactMessage({ From: { EmailAddress: { Address: 'x@y.z', Name: 'x@y.z' } } }).From,
    ).toBe('x@y.z');
  });

  it('drops false/empty fields rather than emitting empty keys', () => {
    const c = compactMessage(message);
    expect(c).not.toHaveProperty('HasAttachments'); // false → dropped
    expect(Object.values(c)).not.toContain(undefined);
  });

  it('keeps IsRead even when false, because false is the meaningful value', () => {
    expect(compactMessage(message).IsRead).toBe(false);
  });

  it('adds correspondence detail only in the full view', () => {
    const c = compactMessage(message);
    const f = fullMessage(message);
    expect(c).not.toHaveProperty('Cc');
    expect(f.Cc).toEqual(['carol@example.com']);
    expect(f.ConversationId).toBe('c1');
    expect(f.WebLink).toBe('https://outlook.office.com/x');
  });

  it('falls back to Sender when From is absent', () => {
    expect(
      compactMessage({ Sender: { EmailAddress: { Address: 's@x.y' } } }).From,
    ).toBe('s@x.y');
  });
});

describe('recipient cap', () => {
  const crowd = Array.from({ length: 1132 }, (_, i) => ({ EmailAddress: { Address: `u${i}@example.com` } }));

  it('keeps at most MAX_RECIPIENTS of To and Cc, with the full count', () => {
    expect(MAX_RECIPIENTS).toBe(20);
    const f = fullMessage({ ...message, ToRecipients: crowd, CcRecipients: crowd.slice(0, 21) });
    expect(f.To).toHaveLength(20);
    expect(f.ToCount).toBe(1132);
    expect(f.ToTruncated).toBe(true);
    expect(f.Cc).toHaveLength(20);
    expect(f.CcCount).toBe(21);
    expect(f.CcTruncated).toBe(true);
    const c = compactMessage({ ...message, ToRecipients: crowd });
    expect(c.To).toHaveLength(20);
    expect(c.ToCount).toBe(1132);
  });

  it('adds no count to a list of exactly MAX_RECIPIENTS', () => {
    const f = fullMessage({ ...message, ToRecipients: crowd.slice(0, 20) });
    expect(f.To).toHaveLength(20);
    expect(f).not.toHaveProperty('ToCount');
    expect(f).not.toHaveProperty('ToTruncated');
  });
});

describe('event projection', () => {
  const event = {
    Id: 'e1',
    Subject: 'Standup',
    Start: { DateTime: '2026-09-22T15:00:00', TimeZone: 'Eastern Standard Time' },
    End: { DateTime: '2026-09-22T15:15:00', TimeZone: 'Eastern Standard Time' },
    Location: { DisplayName: 'Room 2' },
    Organizer: { EmailAddress: { Name: 'Dana', Address: 'dana@example.com' } },
    Attendees: [
      { EmailAddress: { Address: 'e@x.y' }, Status: { Response: 'Accepted' } },
    ],
    IsAllDay: false,
  };

  it('flattens the nested date/location shape', () => {
    const c = compactEvent(event);
    expect(c.Start).toBe('2026-09-22T15:00:00');
    expect(c.TimeZone).toBe('Eastern Standard Time');
    expect(c.Location).toBe('Room 2');
    expect(c.Organizer).toBe('Dana <dana@example.com>');
  });

  it('adds attendee responses only in the full view', () => {
    expect(compactEvent(event)).not.toHaveProperty('Attendees');
    expect(fullEvent(event).Attendees).toEqual([{ Who: 'e@x.y', Response: 'Accepted' }]);
  });

  // Live 2026-10-10 (clearing OOO days): without these an agent could not
  // tell its own appointment from an invite it had not answered, or one
  // occurrence from the whole series.
  const occurrence = {
    ...event,
    ResponseStatus: { Response: 'NotResponded', Time: '0001-01-01T00:00:00Z' },
    ResponseRequested: true,
    IsOrganizer: false,
    IsCancelled: false,
    IsAllDay: false,
    Type: 'Occurrence',
    SeriesMasterId: 'master-1',
    ShowAs: 'Tentative',
  };

  it('carries the decision fields in the compact view, false values included', () => {
    expect(compactEvent(occurrence)).toMatchObject({
      MyResponse: 'NotResponded',
      ResponseRequested: true,
      IsOrganizer: false,
      IsCancelled: false,
      IsAllDay: false,
      Type: 'Occurrence',
      ShowAs: 'Tentative',
    });
    // The series link is for the full view; compact stays compact.
    expect(compactEvent(occurrence)).not.toHaveProperty('SeriesMasterId');
  });

  it('carries the decision fields and the series link in the full view', () => {
    expect(fullEvent(occurrence)).toMatchObject({
      MyResponse: 'NotResponded',
      ResponseRequested: true,
      IsOrganizer: false,
      IsCancelled: false,
      IsAllDay: false,
      Type: 'Occurrence',
      SeriesMasterId: 'master-1',
      ShowAs: 'Tentative',
    });
  });

  it('marks your own meeting as organizer and leaves out what Outlook did not send', () => {
    const own = compactEvent({ Id: 'e2', IsOrganizer: true, ResponseStatus: { Response: 'Organizer' } });
    expect(own).toMatchObject({ IsOrganizer: true, MyResponse: 'Organizer' });
    expect(own).not.toHaveProperty('Type');
    expect(own).not.toHaveProperty('ResponseRequested');
  });
});

describe('folder projection', () => {
  it('keeps a zero count, which is real information', () => {
    expect(compactFolder({ Id: 'f', DisplayName: 'Archive', UnreadItemCount: 0, TotalItemCount: 0 }))
      .toEqual({ Id: 'f', Name: 'Archive', Unread: 0, Total: 0 });
  });

  it('says how many child folders there are, so the model knows to drill in', () => {
    // outlook_list_folders lists one level. Without this count a nested
    // folder (Inbox/Receipts) is invisible: nothing says it exists.
    expect(
      compactFolder({ Id: 'i', DisplayName: 'Inbox', UnreadItemCount: 1, TotalItemCount: 9, ChildFolderCount: 2 }),
    ).toEqual({ Id: 'i', Name: 'Inbox', Unread: 1, Total: 9, Children: 2 });
  });
});

describe('collection envelope', () => {
  it('reports a count and preserves the paging link', () => {
    const out = projectCollection(
      { value: [message], '@odata.nextLink': 'https://outlook.office.com/next' },
      compactMessage,
      'message',
    );
    expect(out.count).toBe(1);
    expect(out.nextLink).toBe('https://outlook.office.com/next');
  });

  it('omits nextLink on the last page', () => {
    expect(projectCollection({ value: [] }, compactMessage, 'message')).not.toHaveProperty(
      'nextLink',
    );
  });

  it('tolerates a missing value array rather than throwing', () => {
    expect(projectCollection({}, compactMessage, 'message').count).toBe(0);
  });

  it('falls back to the raw record when a projection throws', () => {
    // A record with holes is indistinguishable from "there was nothing there",
    // so projectOrRaw hands back everything instead.
    const boom = () => {
      throw new Error('unexpected shape');
    };
    const out = projectCollection({ value: [{ Id: 'x' }] }, boom, 'message');
    expect(out.items).toEqual([{ Id: 'x' }]);
  });
});

describe('VIEWS', () => {
  it('is the one shared view list every read tool offers', () => {
    expect(VIEWS).toEqual(['compact', 'full', 'raw']);
  });
});
