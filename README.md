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

## Tools

**Read** — `outlook_list_folders`, `outlook_list_messages`,
`outlook_get_message`, `outlook_get_unread`, `outlook_list_attachments`,
`outlook_get_attachment`, `outlook_list_events`, `outlook_get_event`,
`outlook_list_calendars`, `outlook_find_meeting_times`,
`outlook_get_schedule`, `outlook_get_profile`,
`outlook_get_mailbox_settings`, `outlook_list_contacts`, `outlook_list_people`,
`outlook_list_categories`, `outlook_list_tasks`

**Write** (ask you to confirm first — see [Confirmations](#confirmations)) —
`outlook_send_mail`, `outlook_create_draft`, `outlook_mark_read`,
`outlook_update_message`, `outlook_move_message`, `outlook_create_event`,
`outlook_update_event`, `outlook_respond_to_invite`, `outlook_reply`.
Two skip the confirmation when nothing reaches anyone else:
`outlook_reply` with `draftOnly: true` (the reply stays in Drafts) and
`outlook_respond_to_invite` with `sendResponse: false` (only your calendar
changes).

**Meetings** — `outlook_find_meeting_times` asks Outlook's Scheduling
Assistant for slots when everyone is free, and `outlook_get_schedule` shows
each person's busy blocks. Each suggested slot carries `createEventArgs`,
which `outlook_create_event` accepts unchanged. `outlook_create_event` and
`outlook_update_event` attach a Microsoft Teams meeting by default and return
its join link; pass `teamsMeeting: false` to leave it off. Only the organizer
can update a meeting. `outlook_respond_to_invite` accepts, tentatively accepts
or declines an invite, by its message id or its event id.

**Diagnostics** — `outlook_healthcheck`

Every read tool takes `view: compact | full | raw`, defaulting to **compact**.
Mutating tools **write nothing** until you confirm (bar the two cases above)
— see [Confirmations](#confirmations). The preview shows exactly what would be sent
(action, method, path and body). (`outlook_create_event` first reads the
mailbox time zone, so its preview can name the zone it would book in; that is
the one read a preview makes.)

Mail and event text is written by other people, so `outlook_list_messages`,
`outlook_get_message`, `outlook_get_unread`, `outlook_get_attachment`,
`outlook_list_events`, `outlook_get_event`, `outlook_find_meeting_times` and
`outlook_get_schedule` wrap every result (all views) in an untrusted-content
envelope — `untrusted_content: true` plus a `note` telling the model to treat
the text as data, never instructions (a `view: raw` record that carries its
own `untrusted_content`/`note` key is nested under `data`, so it cannot
overwrite the envelope).
On a client that cannot show a confirmation prompt, the confirmToken is still
something the model passes back itself, so keep `MCP_CONFIRM_MODE=ask-user`
(the default) and approve each preview in chat — or use a client that asks you
before running non-read-only tools — as the real guard on outbound sends.

## Triage loop

An agent can work through new mail end to end without opening Outlook:

1. **Read the batch** — `outlook_get_unread` returns the unread messages in a
   folder (default inbox), newest first, each with its plain-text body, in one
   call. It does **not** mark anything read. Each item has a `kind`: `mail`,
   `meetingRequest`, `meetingCancelled` or `meetingResponse`; requests and
   cancellations also carry the linked `event` (time, organizer, your current
   response).
2. **Act on each item** —
   - an invite: `outlook_respond_to_invite` with the message id and
     `accept`, `tentative` or `decline` (and an optional comment);
   - a message that needs an answer: `outlook_reply` (`reply`, `replyAll` or
     `forward`), or `draftOnly: true` to leave it in Drafts for you to review;
   - a request for a meeting: `outlook_find_meeting_times` for the attendees,
     then pass the chosen slot's `createEventArgs` to `outlook_create_event`;
   - an attachment to read: `outlook_get_attachment`.
3. **File it** — `outlook_update_message` flags, categorises (names from
   `outlook_list_categories`) and marks read up to 50 messages in one call;
   `outlook_move_message` moves one to another folder.
4. **Mark it handled** — `outlook_mark_read` (or `isRead: true` on
   `outlook_update_message`), so the next `outlook_get_unread` returns only
   what is still new.

Replies, forwards, invite responses and new meetings reach other people, so
each asks you to confirm first.

## Confirmations

Every write asks you to confirm before anything is sent or changed (except a
draft-only reply and a silent invite response — see [Tools](#tools)). A client
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
