# Outlook — inbox triage, replies and meeting-invite responses

**Goal.** Make it easy for an agent to work through new/unread mail end to
end: pull the unread batch with bodies in one call, spot meeting invites,
accept/tentatively-accept/decline them, reply/reply-all/forward, file and
flag what it has handled, and schedule follow-up meetings (find times →
create the invite) — all on the Outlook REST v2.0 API this MCP already uses
(token captured once, plain `fetch` afterwards; no screen reading).

**Tech / conventions.** Same as the repo: `OutlookClient` (`src/client.ts`),
zod 4 schemas, `minifiedResult`/`_untrusted.ts` wrapping for third-party text,
`confirmWrite` + `confirmTokenParam` for every write that sends something to
another person, `writeOrUnknown` for non-idempotent writes, write-verification
warnings where a read-back is cheap. TDD; `npm test` (typecheck + vitest) green
at every commit. Read `src/tools/writes.ts`, `src/tools/mail.ts`,
`src/tools/calendar.ts` and `src/client.ts` before starting any task and match
their style and comment density.

## API facts (Outlook REST v2.0, `https://outlook.office.com/api/v2.0`)

- Meeting invites arrive as messages whose `@odata.type` is
  `#Microsoft.OutlookServices.EventMessage` with `MeetingMessageType`
  (`MeetingRequest`, `MeetingCancelled`, `MeetingAccepted`,
  `MeetingTenativelyAccepted` (sic), `MeetingDeclined`). The linked event:
  `GET /me/messages/{id}?$expand=Microsoft.OutlookServices.EventMessage/Event`
  (property `Event`, with `Id`, `Subject`, `Start`, `End`, `Location`,
  `Organizer`, `ResponseStatus`, `IsCancelled`).
  Filtering the inbox for invites: `$filter=MeetingMessageType eq Microsoft.OutlookServices.MeetingMessageType'MeetingRequest'`
  is NOT reliably supported on `/messages`; instead select
  `MeetingMessageType` and classify client-side.
- Responding: `POST /me/events/{eventId}/accept | tentativelyaccept | decline`
  body `{ "Comment": string, "SendResponse": boolean }` → 202, no body.
  Read-back: `GET /me/events/{eventId}?$select=ResponseStatus` →
  `ResponseStatus.Response` ∈ `Accepted|TentativelyAccepted|Declined|…`.
- Replies: `POST /me/messages/{id}/reply | replyall | forward` body
  `{ "Comment": string, "ToRecipients"?: [...] }` (forward needs
  `ToRecipients`) → 202, sends immediately. Draft variants:
  `POST /me/messages/{id}/createreply | createreplyall | createforward` → 201
  with the draft message (returns its `Id`; the draft can be sent later with
  `POST /me/messages/{draftId}/send`).
- Flag / categories: `PATCH /me/messages/{id}` with
  `{ "Flag": { "FlagStatus": "Flagged"|"Complete"|"NotFlagged" } }` and/or
  `{ "Categories": [string] }`. Master categories:
  `GET /me/outlook/masterCategories`.
- Attachment content: `GET /me/messages/{id}/attachments/{attId}` →
  `#Microsoft.OutlookServices.FileAttachment` with `ContentBytes` (base64),
  `ContentType`, `Name`, `Size`; `ItemAttachment` / `ReferenceAttachment`
  have no bytes.

If any of these turns out wrong against the existing fixtures/clients in the
repo, prefer what the repo already does and note it in the commit body.

## Task 1 — `outlook_get_unread` (triage batch)

**Files:** `src/tools/mail.ts` (+ tests).

Input `{ folder? (default inbox), limit? 1..50 (default 25), sinceHours?,
includeBody? (default true, text only, truncated to `maxBodyChars` default
4000), nextLink? }`. One list call (`IsRead eq false`, newest first, `$select`
incl. `MeetingMessageType`, `@odata.type` handled), then bodies fetched
(concurrency 4) when `includeBody`. Each item: `id, receivedAt, from, to, cc,
subject, importance, hasAttachments, categories, flag, conversationId,
kind ('mail'|'meetingRequest'|'meetingCancelled'|'meetingResponse'),
body (text) / bodyTruncated`. For `meetingRequest`/`meetingCancelled` also
`event: { id, subject, start, end, location, organizer, responseStatus,
isCancelled }` via the `$expand` above. Wrapped as untrusted. Does NOT mark
read. Description tells the agent the processing loop: read → act (reply /
respond to invite / file / flag) → `outlook_mark_read`.

## Task 2 — `outlook_respond_to_invite`

**Files:** `src/tools/writes.ts` (+ tests).

Input `{ messageId? | eventId?` (exactly one; a messageId is resolved to its
event via `$expand`), `response: 'accept'|'tentative'|'decline'`,
`comment?`, `sendResponse? (default true)`, `confirmToken? }`. Uses
`confirmWrite` when `sendResponse` is true (it notifies the organizer),
preview shows subject/time/organizer/response/comment. Refuse with a clear
`McpToolError` when the event is cancelled or the user is the organizer.
`writeOrUnknown`; then read back `ResponseStatus` and warn if it did not
change. Annotations: `destructiveHint: false, idempotentHint: true,
openWorldHint: true`.

## Task 3 — `outlook_reply` (reply / reply-all / forward, send or draft)

**Files:** `src/tools/writes.ts` (+ tests).

Input `{ messageId, mode: 'reply'|'replyAll'|'forward', comment (the new
text), to? (required for forward), draftOnly? (default false),
confirmToken? }`. `draftOnly` → `createreply*`/`createforward` then PATCH the
body/recipients as needed and return the draft id (no confirm needed —
nothing leaves the mailbox). Otherwise `confirmWrite` (preview: mode,
recipients as Outlook will compute them — for reply/replyAll read the
original's From/To/Cc and show them — subject, comment) then the send
endpoint inside `writeOrUnknown` (check Sent Items before retrying).
Annotations `destructiveHint: true` when sending.

## Task 4 — `outlook_update_message` (flag / categories / read state)

**Files:** `src/tools/writes.ts` (+ tests).

Input `{ messageIds: string[] (1..50), flag?: 'flagged'|'complete'|'none',
categories?: string[] (replaces), addCategories?: string[],
removeCategories?: string[], isRead?: boolean }` (at least one change).
add/remove read-modify-write the current categories. Batch with
concurrency 4; per-id result `{ id, ok, error? }` plus a summary. Idempotent,
not destructive. Also add `outlook_list_categories` (master categories,
read-only) in `src/tools/directory.ts`.

## Task 5 — `outlook_get_attachment`

**Files:** `src/tools/mail.ts` (+ tests).

Input `{ messageId, attachmentId }`. FileAttachment → for `image/*` return
MCP image content; for text-like types (`text/*`, json, csv, xml, ics)
return decoded text (cap 200 KB); else return metadata + `contentBytes`
omitted with `hint` that binary content is not inlined (cap 5 MB, never inline
above). Item/Reference attachments → metadata only, explaining why.
Untrusted wrapping for text.

## Task 6 — schedule-a-meeting flow polish

**Files:** `src/tools/calendar.ts`, `src/tools/writes.ts` (+ tests).

`outlook_find_meeting_times` and `outlook_create_event` already exist. Make
the chain agent-friendly without new endpoints: each suggestion returned by
`outlook_find_meeting_times` carries a ready-to-use
`createEventArgs: { start, end, timeZone, attendees, subject? }` matching
`outlook_create_event`'s input exactly, and `outlook_create_event`'s
description points at it. Verify by a test that feeds a suggestion's
`createEventArgs` straight into the create tool's zod schema.

## Task 7 — manifests, docs, skill

`manifest.json` tools list, README tool table + "triage loop" section, the
repo's skill doc (if it lists tools), served-tools/annotation/packaging tests.

## Live verification (orchestrator, with the user — after all tasks)

Unread batch matches the inbox; an invite is detected and its event
resolved; respond to a real invite only with the user's explicit go-ahead;
draft-only reply lands in Drafts; flag + category round-trip; attachment
read.
