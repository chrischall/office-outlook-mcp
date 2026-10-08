# Outlook REST v2.0 — verified request shapes

Everything here was executed live against a real Microsoft 365 mailbox on
**2026-09-20** and returned the status shown. Nothing is inferred from
documentation. No credential, cookie or token is recorded in this file.

Base: `https://outlook.office.com/api/v2.0`

---

## How auth actually works

The signed-in Outlook Web tab lives on **`outlook.cloud.microsoft`** (the new
Outlook Web domain), but the `Authorization: Bearer` header it sends carries a
token whose claims are:

| claim | value |
| --- | --- |
| `aud` | `https://outlook.office.com` |
| `appid` | `9199bf20-a13f-4107-85dc-02114787ef48` ("One Outlook Web") |
| `iss` | `https://sts.windows.net/<tenant>/` |
| lifetime | `exp - iat` ≈ **25 hours** |

`scp` includes `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
`Contacts.ReadWrite`, `Tasks.ReadWrite`, `MailboxSettings.ReadWrite`,
`People.Read`, `User.Read.All`.

**That token authenticates plain server-side requests** to
`outlook.office.com/api/v2.0` — no browser, no bridge, no tab. This is what
makes the repo a *bootstrap* archetype rather than a full-bridge one: the
browser is needed only to mint, never to call.

There is no refresh grant we can drive (the token comes from MSAL's hidden
iframe against `login.microsoftonline.com`), so "refresh" means "capture
again".

### Token capture

`captureRequestHeader` on `authorization@outlook.cloud.microsoft`, racing the
same declaration on `authorization@outlook.office.com` for tenants still on the
legacy host. Capture resolves on the **next** matching request, so the tab must
make one inside the window — a reload guarantees it.

---

## Field casing

**PascalCase.** This is the Outlook REST API, not Microsoft Graph:
`ReceivedDateTime`, `IsRead`, `BodyPreview`, `DisplayName`. A camelCase Graph
snippet returns nothing useful and does *not* error.

---

## Verified endpoints

All returned **200**:

| path | notes |
| --- | --- |
| `/me` | `Id, EmailAddress, DisplayName, Alias, MailboxGuid` |
| `/me/mailfolders` | `Id, DisplayName, ParentFolderId, ChildFolderCount, UnreadItemCount, TotalItemCount, SizeInBytes, IsHidden` |
| `/me/mailfolders/{wellKnown}` | `inbox, drafts, sentitems, deleteditems, archive, junkemail, outbox, clutter` — all eight present |
| `/me/mailfolders/{id}/messages` | folder-scoped listing |
| `/me/messages` | mailbox-wide; required for `$search` |
| `/me/messages/{id}` | single message, includes `Body` |
| `/me/messages/{id}/attachments` | `Id, Name, ContentType, Size` (+ `ContentBytes` unless `$select`ed out) |
| `/me/calendars` | calendar list |
| `/me/events` | does **not** expand recurring series |
| `/me/calendarview?startDateTime=&endDateTime=` | **does** expand recurrences — use this for "what's on my schedule" |
| `/me/contacts` | saved address book |
| `/me/people` | relevance-ranked, drawn from mail traffic |
| `/me/tasks` | Outlook tasks |
| `/me/MailboxSettings` | time zone, working hours, auto-replies |

Query options verified: `$top`, `$skip`, `$select`, `$filter`, `$orderby`,
`$search`, `$count`.

**`$search` cannot be combined with `$filter`** — and in practice not with
`$orderby` either, so the client omits both when searching.

Paging: `@odata.nextLink` is an **absolute URL** on the same origin.

### `Prefer` headers

Both verified live:

- `outlook.body-content-type="text"` — **measured 7,299 bytes HTML → 775 bytes
  text on one real message (9.4×)**; across a 5-message collection 37,638 →
  12,828. Also normalises a collection that otherwise mixes `HTML` and `Text`.
- `outlook.timezone="Eastern Standard Time"` — `Start.TimeZone` comes back as
  the requested zone. Takes **Windows** zone names, not IANA
  (`America/New_York` is rejected).

Multiple values combine in one comma-separated header; confirmed working.

---

## Scheduling assistant (read-only POSTs)

Verified live **2026-10-08**, both **200**. They are POSTs but change
nothing, so the client sends them through `post`, not `write`.

### `POST /me/findmeetingtimes`

```json
{
  "Attendees": [{ "Type": "Required", "EmailAddress": { "Address": "a@x" } }],
  "TimeConstraint": {
    "ActivityDomain": "Work",
    "Timeslots": [{
      "Start": { "DateTime": "2026-10-12T09:00:00", "TimeZone": "Eastern Standard Time" },
      "End":   { "DateTime": "2026-10-13T17:00:00", "TimeZone": "Eastern Standard Time" }
    }]
  },
  "MeetingDuration": "PT45M",
  "MaxCandidates": 3,
  "ReturnSuggestionReasons": true
}
```

Returns `EmptySuggestionsReason` plus `MeetingTimeSuggestions[]`, each with
`Confidence`, `OrganizerAvailability`, `SuggestionReason`,
`MeetingTimeSlot.{Start,End}` and `AttendeeAvailability[]`
(`{Availability, Attendee.EmailAddress.Address}`).

- **Slots come back in UTC** regardless of the request's `TimeZone` unless
  `Prefer: outlook.timezone="…"` is sent. With it they are in that zone.
- Listing the signed-in user as an attendee returns an **empty**
  `AttendeeAvailability`: they are the organizer, and appear only as
  `OrganizerAvailability`.
- `PT30M` and `PT45M` both accepted.

### `POST /me/calendar/getschedule`

```json
{
  "Schedules": ["a@x"],
  "StartTime": { "DateTime": "2026-10-12T09:00:00", "TimeZone": "Eastern Standard Time" },
  "EndTime":   { "DateTime": "2026-10-12T17:00:00", "TimeZone": "Eastern Standard Time" },
  "AvailabilityViewInterval": 30
}
```

Returns `value[]` of `{ScheduleId, AvailabilityView, ScheduleItems[]}`, with
items carrying `Status, Subject, Location, IsMeeting, IsRecurring, IsPrivate,
Start, End`. `AvailabilityView` is one digit per interval (`0` free,
`2` busy). The same `Prefer` timezone rule applies.

An address that does not resolve still returns 200, as an entry with
`Error: {Message: "MailRecipientNotFoundException…", ResponseCode: "5009"}`
and no items. That must not be reported as "free".

---

## Writes

Each route was confirmed to exist and be authorised by POSTing a deliberately
invalid body and receiving **400**, not 401/403/404:

| route | error on invalid body |
| --- | --- |
| `POST /me/sendmail` | `RequestBodyRead` |
| `POST /me/messages` | `UnableToDeserializePostBody` |
| `POST /me/events` | `UnableToDeserializePostBody` |
| `POST /me/contacts` | `UnableToDeserializePostBody` |

**No mail write was executed during the build** — no mail was sent and no
message was mutated (calendar exceptions below). The payload shapes the client sends are therefore
*shapes*, not round-tripped captures, and every mutating tool asks for
confirmation first — a prompt where the client supports one, otherwise a
preview plus a single-use confirmToken (`MCP_CONFIRM_MODE`).

`outlook_mark_read` verifies by re-reading `IsRead` — the field the write
actually requested — rather than trusting the status code, and reports a
warning when the value did not move.

Note `POST /me/messages/{id}/move` assigns the message a **new `Id`**; the old
one stops resolving.

### Teams meetings (executed live 2026-10-08)

Unlike the routes above, these were **round-tripped**, on two test events with
no attendees (so no invitations went out), and both were deleted afterwards:

- `POST /me/events` with `"IsOnlineMeeting": true, "OnlineMeetingProvider":
  "TeamsForBusiness"` returns the created event with `OnlineMeeting.JoinUrl`
  already populated, and Outlook sets `Location` to "Microsoft Teams Meeting".
- `PATCH /me/events/{id}` with the same two fields adds a Teams meeting to an
  existing event. The re-read shows the join link.
- `PATCH` with `Subject` and `Start`/`End` (as `DateTimeTimeZone`) persisted
  as sent.
- `DELETE /me/events/{id}` removed the event.

The Teams link is at **`OnlineMeeting.JoinUrl`**. The older `OnlineMeetingUrl`
is `""` on every Teams meeting read (73 of 73 over two weeks). An event with
no online meeting has `OnlineMeeting: {}`, `IsOnlineMeeting: false` and
`OnlineMeetingProvider: "Unknown"`. `$select=OnlineMeeting` works.

`IsOrganizer: false` marks a meeting someone else organizes. An attendee's
PATCH would change only their own copy, so `outlook_update_event` refuses it.

---

## Errors

| signal | meaning |
| --- | --- |
| `401`, **empty body**, `WWW-Authenticate: Bearer client_id="00000002-0000-0ff1-ce00-000000000000"` | token expired or revoked → re-capture |
| `400` + `error.code` | route exists, payload wrong |
| `400 RequestBroker--ParseUri` | the **path** is not a route at all (an unknown segment gives 400, not 404) |
| `404` | valid route, bad id |
| `500` on `/owa/0/service.svc` | that is OWA's internal endpoint, not this API |
