# office-outlook-mcp

MCP server for **Outlook / Microsoft 365** — mail, folders, calendar, contacts
and tasks, with confirmation-gated sending.

> This project was developed and is maintained by AI (Claude Code). Use at your
> own discretion.

## How it authenticates

Outlook Web holds its access token in memory and sends it as
`Authorization: Bearer …`. This server snapshots that header from your
signed-in browser tab through the [fetchproxy](https://github.com/chrischall/fetchproxy)
bridge, then talks to `https://outlook.office.com/api/v2.0` with ordinary
server-side requests.

**The browser is needed only to mint the token (~25h), never to make a call.**
That is the whole design: no Azure app registration, no admin consent, and no
browser on the request path.

Tokens are cached at `~/.office-outlook-mcp/token.json` (0600), so a restart
does not re-capture.

## Install

```sh
npm i -g @chrischall/office-outlook-mcp
```

Requires the **ContextMint Bridge** browser extension, installed from
[its releases](https://github.com/nullnet-app/contextmint-bridge/releases):
in Chrome, unzip the chrome build and load it unpacked
(`chrome://extensions` → Developer mode → Load unpacked). Safari is not
available yet — it will ship inside the ContextMint app, which has no public
download — so use Chrome for now.

ContextMint Bridge is the fetchproxy browser extension under its new name, from
the same maintainer — fetchproxy's own README
([#extension](https://github.com/chrischall/fetchproxy#extension)) points to it.
Its source is public at
[nullnet-app/contextmint-bridge](https://github.com/nullnet-app/contextmint-bridge):
build it yourself, or check a release zip against the `.sha256` file published
beside it (`shasum -a 256 -c contextmint-bridge-chrome-<version>.zip.sha256`).

The extension, `fpx` and this server are compatible by **fetchproxy protocol
number** (currently protocol 4), not by matching package versions. Keep the
bridge and `fpx` current — a mismatch fails with an error naming both protocol
numbers. Also install `@fetchproxy/cli`:

```sh
npm i -g @fetchproxy/cli
```

The first capture prints a 6-digit pair code to approve in the ContextMint
Bridge popup; the grant persists.

### Install in opencode

opencode reads MCP servers from `opencode.json` (project) or
`~/.config/opencode/opencode.json` (global):

```json
{
  "mcp": {
    "servers": {
      "outlook": {
        "type": "local",
        "command": ["npx", "-y", "@chrischall/office-outlook-mcp"]
      }
    }
  }
}
```

That one file serves **opencode 2 and current opencode 1** — verified live on
2.0.11, 1.18.31 and 1.18.28, which all connect from it.

Two things worth knowing, both measured rather than assumed:

- **Never carry both config shapes in one file.** Older opencode 1 wants the
  servers directly under `mcp` (`"mcp": { "outlook": { … } }`), and the
  temptation is to write both so either version finds one. Do not: given a
  sibling key beside `servers`, opencode 2 parses the file — it still shows up
  in `opencode debug config` — and then reports *"No MCP servers configured"*.
  No error, no warning, every server silently gone.
- **An opencode 1 old enough to reject `servers` says so loudly**
  (`Configuration is invalid at …`), so if you see that, switch that machine to
  the flat shape rather than combining them.

opencode 2 talks to a background service, so `opencode reload` before
`opencode mcp list` after editing config.

### Configuration

Everything is optional — with nothing set, the server captures from the browser.

| variable | purpose |
| --- | --- |
| `OUTLOOK_ACCESS_TOKEN` | Pre-obtained bearer token. Overrides the bridge. Power users / CI only; expires in ~25h and **cannot be auto-refreshed**. |
| `OUTLOOK_DISABLE_FETCHPROXY` | `1` to disable browser capture, requiring `OUTLOOK_ACCESS_TOKEN`. |
| `OUTLOOK_API_BASE` | Override the REST base URL. Defaults to `https://outlook.office.com/api/v2.0`. |
| `OUTLOOK_WS_PORT` | fetchproxy concentrator port. Defaults to `37149`, the fleet-wide shared port. |
| `OUTLOOK_CAPTURE_TIMEOUT` | Seconds to wait for the tab to make a readable request. Bounds the WHOLE capture, both declared hosts included. Defaults to `30`. |
| `OUTLOOK_TOKEN_CACHE` | `false` to stop caching the token between runs. |
| `OUTLOOK_TOKEN_FILE` | Path to the cached token. Defaults to `~/.office-outlook-mcp/token.json`. |
| `OUTLOOK_OUTPUT_DIR` | Absolute directory `outlook_export_messages` and `outlook_export_events` write into. Defaults to `~/Downloads/office-outlook-mcp`. When set, a per-call `dest` must be inside it. See [Bulk export](#bulk-export). |

## Tools

**Read** — `outlook_list_folders`, `outlook_list_messages`,
`outlook_get_message`, `outlook_list_unread_messages`, `outlook_list_attachments`,
`outlook_get_attachment`, `outlook_list_events`, `outlook_get_event`,
`outlook_list_calendars`, `outlook_find_meeting_times`,
`outlook_get_schedule`, `outlook_get_profile`,
`outlook_get_mailbox_settings`, `outlook_list_contacts`, `outlook_list_people`,
`outlook_list_categories`, `outlook_list_tasks`

**Write** (ask you to confirm first — see [Confirmations](#confirmations)) —
`outlook_send_message`, `outlook_create_draft`, `outlook_mark_read`,
`outlook_update_message`, `outlook_move_message`, `outlook_create_event`,
`outlook_update_event`, `outlook_respond_to_invite`, `outlook_delete_event`,
`outlook_reply`.
Some skip the confirmation because nothing reaches anyone else:
`outlook_create_draft` (the draft stays in Drafts until you send it from
Outlook), `outlook_reply` with `draftOnly: true` (the reply stays in Drafts),
`outlook_respond_to_invite` with `sendResponse: false`, and
`outlook_delete_event` on your own appointment with no attendees or on a
meeting the organizer already cancelled (in each, only your calendar
changes).

**Meetings** — `outlook_find_meeting_times` asks Outlook's Scheduling
Assistant for slots when everyone is free, and `outlook_get_schedule` shows
each person's busy blocks. Each suggested slot carries `createEventArgs`,
which `outlook_create_event` accepts unchanged. `outlook_create_event`
attaches a Microsoft Teams meeting by default and returns its join link; pass
`teamsMeeting: false` to leave it off. `outlook_update_event` never changes how
people join a meeting unless asked: pass `teamsMeeting: true` to add a Teams
meeting to an event that lacks one, and the confirmation says so in plain
words. Only the organizer
can update a meeting. `outlook_respond_to_invite` accepts, tentatively accepts
or declines an invite, by its message id or its event id.
`outlook_delete_event` deletes an event you organize — one occurrence by the
occurrence's id, or the whole series by its series master's id — and sends
any attendees a cancellation; for an invite you received it refuses and points
you at declining instead. `outlook_list_events` and `outlook_get_event` show
what that decision needs in every view: `MyResponse`, `ResponseRequested`,
`IsOrganizer`, `IsCancelled`, `Type` (SingleInstance, Occurrence, Exception or
SeriesMaster), `ShowAs` and `IsAllDay`, plus `SeriesMasterId` in the full view.

**Export** — `outlook_export_messages` (mail) and `outlook_export_events`
(calendar) write JSONL or Markdown files, optionally with attachments, and
return only where they are; see [Bulk export](#bulk-export).

**Diagnostics** — `outlook_healthcheck`

Every read tool takes `view: compact | full | raw`, defaulting to **compact**.
Mutating tools **write nothing** until you confirm (bar the cases above)
— see [Confirmations](#confirmations). The preview shows exactly what would be sent
(action, method, path and body). (`outlook_create_event` first reads the
mailbox time zone, so its preview can name the zone it would book in; that is
the one read a preview makes.)

Mail and event text is written by other people, so `outlook_list_messages`,
`outlook_get_message`, `outlook_list_unread_messages`, `outlook_get_attachment`,
`outlook_list_events`, `outlook_get_event`, `outlook_find_meeting_times`,
`outlook_get_schedule`, `outlook_export_messages` and `outlook_export_events`
wrap every result (all views) in an untrusted-content
envelope — `untrusted_content: true` plus a `note` telling the model to treat
the text as data, never instructions (a `view: raw` record that carries its
own `untrusted_content`/`note` key is nested under `data`, so it cannot
overwrite the envelope). `outlook_reply` returns a sent reply's subject — the
original sender's text — inside the same envelope.
On a client that cannot show a confirmation prompt, the confirmToken is still
something the model passes back itself, so keep `MCP_CONFIRM_MODE=ask-user`
(the default) and approve each preview in chat — or use a client that asks you
before running non-read-only tools — as the real guard on outbound sends.

## Triage loop

An agent can work through new mail end to end without opening Outlook:

1. **Read the batch** — `outlook_list_unread_messages` returns the unread messages in a
   folder (default inbox), newest first, each with its plain-text body, in one
   call. It does **not** mark anything read. Each item has a `kind`: `mail`,
   `meetingRequest`, `meetingCancelled` or `meetingResponse`; requests and
   cancellations also carry the linked `event` (time, organizer, your current
   response) — or, when the meeting is no longer on your calendar (already
   declined, cancelled or deleted), a `hint` saying there is nothing to respond
   to. `to` and `cc` list at most 20 addresses, with `toCount`/`ccCount` when
   there are more.
2. **Act on each item** —
   - an invite: `outlook_respond_to_invite` with the message id and
     `accept`, `tentative` or `decline` (and an optional comment). Once you
     respond, Outlook moves the invite message to Deleted Items under a new
     id, so an answered invite is already filed — skip steps 3 and 4 for it
     (its old id would 404);
   - a message that needs an answer: `outlook_reply` (`reply`, `replyAll` or
     `forward`), or `draftOnly: true` to leave it in Drafts for you to review;
   - a request for a meeting: `outlook_find_meeting_times` for the attendees,
     then pass the chosen slot's `createEventArgs` to `outlook_create_event`;
   - an attachment to read: `outlook_get_attachment`.
3. **File it** — `outlook_update_message` flags, categorises (names from
   `outlook_list_categories`) and marks read up to 50 messages in one call;
   `outlook_move_message` moves one to another folder.
4. **Mark it handled** — `outlook_mark_read` (or `isRead: true` on
   `outlook_update_message`) on the plain mail you handled, so the next
   `outlook_list_unread_messages` returns only what is still new. Answered invites are
   already out of the inbox (see step 2).

Replies, forwards, invite responses and new meetings reach other people, so
each asks you to confirm first — except a reply saved with `draftOnly: true`
(it stays in Drafts until you send it) and an invite response with
`sendResponse: false` (only your calendar changes; the organizer is not told).

## Paging from a time

`outlook_list_unread_messages` and `outlook_list_messages` also page **oldest first**
from a point in time: pass `from` as an ISO 8601 UTC time
(`2026-10-01T00:00:00Z`). The result carries `next_from` — pass it back
unchanged as `from` for the next page — and `has_more`. To backfill, start
from an old time and keep passing `next_from` until `has_more` is false; then
keep polling with the last `next_from` to pick up new mail (an empty page
returns the same position). Messages that share a timestamp are never skipped
or repeated across pages, and marking mail read between pages loses nothing.
`from` cannot be combined with `nextLink`, `sinceHours`, `search` or `skip`.

## Bulk export

`outlook_export_messages` exports mail and `outlook_export_events` exports
the calendar in bulk, thousands of records at a time, and none of their text
enters the conversation. The server pages the mailbox (or calendar) and writes
the records to files. The tool returns only a small manifest: where the files
are, their sizes, the record and attachment counts, the date span, and
`has_more` plus a `cursor`. Have the agent read the files with shell tools
(`jq`, `grep`, `rg`, `head`) rather than opening them whole.

- **Mail arguments:** `folder` (a well-known name such as `inbox` or
  `sentitems`, or a folder id; `all` exports the whole mailbox; default
  `inbox`), `since` and `until` (ISO 8601), `search` (full-text, spans the
  mailbox). Mail is exported oldest first.
- **Event arguments:** `since` and `until` (ISO 8601, both required, at most
  two years apart) and `calendar` (a calendar id; default, the default
  calendar). Events come from the calendar view, so a recurring series is
  expanded into its individual occurrences. Times are wall-clock times in the
  mailbox's time zone, named on each record as `time_zone`.
- **Both:** `format` (`jsonl`, the default, or `markdown`), `attachments`
  (default `false`), `dest`, `cursor`, `max_items` (default 5000, max 50000)
  and `budget_seconds` (default 90, max 600).
- **Files:** `part-0001.jsonl` holds one JSON record per message (id,
  conversation id, subject, from, to/cc/bcc, dates, flags, categories, web link
  and the full plain-text body). Its first line is an `_export` header that
  names the source, account, arguments and an untrusted-content note.
  An event record carries the id, iCal UID, type and series id, subject,
  start, end, time zone, location, organizer, attendees with their responses,
  your response, show-as, categories, the join link and the full plain-text
  body.
  `format: "markdown"` writes `part-0001.md`, a readable transcript (one
  section per message or event), instead.
  `manifest.json` lists the counts and any errors, and the files: on disk
  every part's files (it is rewritten as each part finishes); in a blob zip
  the files in that zip.
- **Attachments:** with `attachments: true`, each file attachment of up to
  50 MB is saved as `attachments/<short id>/<name>` (the name sanitized), and
  the record lists its attachments, each with its relative `path`, or with
  `skipped` and the reason. Attached Outlook items (a forwarded email or
  event), cloud-file links (OneDrive/SharePoint), larger files and failed
  downloads are skipped. One bad attachment never fails the export.
- **Resuming:** each call stops at `max_items` or `budget_seconds`, whichever
  comes first. While `has_more` is true, pass `cursor` back unchanged. The
  next call continues the same export as `part-0002`, `part-0003` and so on.
  With attachments on, the budget is checked between small groups of
  records, so a stop can fall inside a page; the cursor resumes at the next
  record. If Outlook refuses the credential (including while downloading an
  attachment), or keeps returning 429 after three
  Retry-After waits, the call keeps what it already wrote and returns the
  cursor. Re-authenticate, then resume.

Where the files go depends on how the server runs:

| mode | when | `location` in the result |
| --- | --- | --- |
| `disk` | any local (stdio) install, the default | an **absolute local directory**, `<dir>/outlook-mail-<YYYYMMDD-HHmmss>/` (or `outlook-events-…`); each file entry carries its absolute `path` |
| `blob` | hosted on mcp-host, which sets both `MCP_BLOB_BASE_URL` and `MCP_BLOB_SIGNING_KEY` | a **signed download link** to one zip per call (`part-NNNN.zip`, which includes `manifest.json`), valid for **1 hour** (`url_expires_at`) |

On disk, `<dir>` is the per-call `dest`, else `OUTLOOK_OUTPUT_DIR`, else
`~/Downloads/office-outlook-mcp`. A `dest` must be inside `OUTLOOK_OUTPUT_DIR`
when that is set. Otherwise it must be inside `~/Downloads`, `~/Documents` or
`~/Desktop`. A resumed cursor's directory is checked against the same roots
again.

## Confirmations

Every write asks you to confirm before anything is sent or changed (except a
new draft, a draft-only reply and a silent invite response — see
[Tools](#tools)). A client
that can show a confirmation prompt (Claude Code) shows one. On a client that
cannot, the first call does nothing and returns a `confirmation-required`
preview plus a `confirmToken`; only a repeat call with that token performs the
write. The token is single-use, expires, and is bound to the exact arguments —
change anything between the two calls and it is refused (`DRAFT_CHANGED`) with
a fresh preview.

| variable | default | |
|---|---|---|
| `MCP_CONFIRM_MODE` | `ask-user` | What a write does on a client that cannot show a confirmation prompt (claude.ai, Claude Desktop). `ask-user`: two steps — the first call does nothing and returns a preview plus a token, and the model must get your approval in chat before calling again with it. `auto`: the same two steps, but the model may use the token after reviewing the preview itself. `refuse`: writes are refused on such clients. A client that can show prompts (Claude Code) always gets the real prompt, unless `MCP_CONFIRM_ELICITATION=off`. An unrecognised value is treated as `refuse`. |
| `MCP_CONFIRM_ELICITATION` | `on` | `off` never shows a confirmation prompt, so every client gets the `MCP_CONFIRM_MODE` behaviour. Set it for a client that says it can show prompts but never does (the write hangs — opencode 2.0.x). Any other value is treated as `on`, with a warning on stderr. |
| `MCP_CONFIRM_TTL_SECONDS` | `600` | How long a token stays valid. |
| `MCP_CONFIRM_SECRET` | random per process | Signing key; set it only if tokens must survive a server restart. On mcp-host the host supplies a stable per-child key (`MCP_HOST_CONFIRM_SECRET`) and spent tokens are recorded under `MCP_DATA_DIR`, so an approval survives an idle restart. |

## Things worth knowing

- Fields are **PascalCase** (`ReceivedDateTime`, `IsRead`). This is the Outlook
  REST API, not Graph — a Graph snippet silently returns nothing.
- `outlook_get_message` returns a **plain-text** body by default, measured 9.4×
  smaller than the HTML; pass `html: true` for the original.
- `outlook_list_events` uses the calendar *view*, which expands recurring
  series. `/me/events` does not, which is why it is not exposed as-is.
- Time zones are **Windows** names (`Eastern Standard Time`), not IANA.
- `search` and `unreadOnly` cannot be combined; the API rejects `$search`
  alongside `$filter`, so the tool refuses before making a doomed request.
- Moving a message assigns it a **new id**.
- `outlook_list_events` returns times in the **mailbox's own time zone** unless
  `timeZone` overrides it. The API itself defaults to UTC, which silently reads
  as a four-hour error on an Eastern mailbox.
- Contact phone fields are `MobilePhone1`, not Graph's `MobilePhone`. The v2.0
  Contact type rejects the Graph name with a 400.

## The lightweight alternative

If you only need Outlook access from Claude Code on this machine, the
`skills/outlook-fpx` access skill in this repo does the same reads with `fpx` +
`curl` and no server at all. The MCP earns its keep when you want typed tools,
confirmation-gated writes, or reach from claude.ai.

## Development

```sh
npm install
npm run build
npm test              # typecheck + suite
npm run test:coverage # CI's gate
```

Request shapes are pinned in [`docs/OUTLOOK-API.md`](docs/OUTLOOK-API.md), all
verified against a live mailbox.

## License

MIT
