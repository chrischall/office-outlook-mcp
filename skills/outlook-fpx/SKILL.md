---
name: outlook-fpx
description: Read and act on the user's Outlook / Microsoft 365 mailbox — mail, folders, calendar, contacts, tasks — via the Outlook REST API, using a bearer token lifted from their signed-in browser tab. Use when asked about their Outlook email, calendar or meetings.
---

# Outlook (Microsoft 365) access

Talks to `https://outlook.office.com/api/v2.0` with a bearer token captured
from the signed-in Outlook Web tab. **The browser is needed only to mint the
token** (~25h lifetime) — every actual request is a plain server-side `curl`.

## Setup (once)

```sh
npm i -g @fetchproxy/cli                      # provides `fpx`
fpx profile add outlook --domain outlook.cloud.microsoft \
                        --domain outlook.office.com
fpx profile declare outlook \
  --capture-header authorization@outlook.cloud.microsoft \
  --capture-header authorization@outlook.office.com \
  --allow-download
```

Declare **all** scopes before the first pairing — widening them later forces the
user to re-approve. The first `fpx` call prints a 6-digit pair code to approve in
the Transporter extension popup; the grant then persists.

Requires the **Transporter** Chrome extension, and the extension and CLI must be
on the **same major version** — a 3.x extension against a 2.x CLI fails by
binding the port and silently never connecting.

## Use

```sh
source references/outlook-env.sh
outlook_token_refresh        # once per ~25h; needs an open, signed-in Outlook tab
outlook_get '/me' | jq .
```

- `outlook_get <path> [curl args]` — GET, body on stdout
- `outlook_get_text <path>` — same plus `Prefer: outlook.body-content-type="text"`; **use this for anything you will read** (measured 9.4× smaller than HTML)
- `outlook_post <path> <json>` / `outlook_patch <path> <json>`
- `outlook_token_expiry` — prints time remaining, never the token

Paths are relative to `/api/v2.0`. Exit code **4** means non-2xx; a **401**
prints `run: outlook_token_refresh`.

Common starting points — full verified set, including calendar, contacts, tasks
and write payloads, in **`references/api-recipes.md`**:

```sh
outlook_get '/me/mailfolders/inbox?$select=UnreadItemCount,TotalItemCount'
outlook_get '/me/mailfolders/inbox/messages?$top=20&$orderby=ReceivedDateTime%20desc&$select=Id,Subject,From,ReceivedDateTime,IsRead'
outlook_get '/me/calendarview?startDateTime=2026-09-20T00:00:00Z&endDateTime=2026-09-27T00:00:00Z&$select=Subject,Start,End,Location'
```

## Rules that bite

- **Fields are PascalCase** (`ReceivedDateTime`, `IsRead`). This is the Outlook
  REST API, *not* Graph — a camelCase Graph snippet silently returns nothing.
- **Always pass `$select`.** The default projection is ~40 fields per message
  including the full body.
- **`$search` and `$filter` cannot be combined.**
- **`/me/events` does not expand recurring series** — use `/me/calendarview`
  with a start/end window for "what's on my schedule".
- **`Prefer: outlook.timezone` takes Windows zone names** (`Eastern Standard
  Time`), not IANA (`America/New_York`).
- **`@odata.nextLink` is an absolute URL** — pass it to `curl -K ~/.outlook-fpx/curlrc`
  directly, not to `outlook_get`, which would prefix the base again.
- **`/owa/0/service.svc` is not this API** and returns 500. Use `/api/v2.0`.
- **A 2xx is not proof a write persisted** — re-read the resource to confirm.
- Never name a shell local `path` or `status`: in zsh `path` is tied to `PATH`
  (assigning it wipes `PATH` mid-function) and `status` is read-only. Both are
  harmless in bash, so this only breaks for zsh users.

## Privacy

This reads the user's real mailbox. Surface only what was asked for; do not dump
message bodies, recipients or attachments into the transcript incidentally. The
token lives in `~/.outlook-fpx/curlrc` (0600, its own directory — deliberately
not shared with any MCP's state) and is passed to `curl` via `-K` so it never
appears in `argv` or `ps` output.
