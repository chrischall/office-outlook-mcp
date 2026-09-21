# office-outlook-mcp

MCP server for **Outlook / Microsoft 365** — mail, folders, calendar, contacts
and tasks, with confirm-gated sending.

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

Requires the **Transporter** Chrome extension and `@fetchproxy/cli`, on a
matching major version:

```sh
npm i -g @fetchproxy/cli
```

The first capture prints a 6-digit pair code to approve in the Transporter
popup; the grant persists.

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
`outlook_get_message`, `outlook_list_attachments`, `outlook_list_events`,
`outlook_get_event`, `outlook_list_calendars`, `outlook_get_profile`,
`outlook_get_mailbox_settings`, `outlook_list_contacts`, `outlook_list_people`,
`outlook_list_tasks`

**Write** (all require `confirm: true`) — `outlook_send_mail`,
`outlook_create_draft`, `outlook_mark_read`, `outlook_move_message`,
`outlook_create_event`

**Diagnostics** — `outlook_healthcheck`

Every read tool takes `view: compact | full | raw`, defaulting to **compact**.
Mutating tools **write nothing** without `confirm: true` — they return a
dry-run preview of exactly what would be sent. (`outlook_create_event` first
reads the mailbox time zone, so its preview can name the zone it would book
in; that is the one read a dry run makes.)

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
confirm-gated writes, or reach from claude.ai.

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
