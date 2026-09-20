import { describe, expect, it } from 'vitest';
import {
  compactEvent,
  compactFolder,
  compactMessage,
  fullEvent,
  fullMessage,
  projectCollection,
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
});

describe('folder projection', () => {
  it('keeps a zero count, which is real information', () => {
    expect(compactFolder({ Id: 'f', DisplayName: 'Archive', UnreadItemCount: 0, TotalItemCount: 0 }))
      .toEqual({ Id: 'f', Name: 'Archive', Unread: 0, Total: 0 });
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
