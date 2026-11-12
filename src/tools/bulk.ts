/**
 * `outlook_move_messages` — move every message in a folder that matches a
 * filter, under ONE confirmation.
 *
 * Built for cleaning up a large mailbox: running an inbox rule over the mail
 * that arrived before the rule existed, or archiving old read mail. One
 * `outlook_move_message` per message (each with its own approval) does not
 * scale to thousands.
 *
 * Every call — the preview AND the confirmed one — lists the matching
 * messages afresh. The confirmation token binds the filter, the destination,
 * markRead, the number matched and the ids that will move, so a mailbox that
 * changed between the preview and the approval gets a fresh preview
 * (`DRAFT_CHANGED`) instead of moving a set the user never saw.
 */
import { createHash } from 'node:crypto';
import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/server';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  currentCallSignal,
  mapWithConcurrency,
  McpToolError,
  minifiedResult,
  reportProgress,
  throwIfCancelled,
  WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import { isCredentialFailure, type OutlookClient, type QueryParams } from '../client.js';
import { workMeta } from './_hints.js';
import { parseTime, utcLiteral } from './_paging.js';
import { unknownOutcome, TIMEOUT_SENTENCE } from './writes.js';

/** How many messages one call moves unless told otherwise. */
export const BULK_MOVE_DEFAULT_MAX = 1000;
/** The most one call ever moves. */
export const BULK_MOVE_HARD_MAX = 5000;
/** Moves in flight at once. */
const MOVE_CONCURRENCY = 4;
/** Rows per listing page. */
const PAGE_SIZE = 100;
/** Messages shown in the preview. */
const SAMPLE_SIZE = 10;
/** Failures (and unknown ids) listed in the result; the rest are only counted. */
const LISTED_FAILURES = 20;
/** The longest subject shown in the preview or a failure. */
const SUBJECT_MAX = 120;
const SELECT = 'Id,Subject,From,ReceivedDateTime,IsRead';

/** The conditions a message must meet to be moved. */
export interface MoveFilter {
  from?: string[];
  receivedBefore?: string;
  receivedAfter?: string;
  isRead?: boolean;
  subjectContains?: string[];
  exceptSubjectContains?: string[];
}

interface Row {
  Id?: string;
  Subject?: string;
  From?: { EmailAddress?: { Address?: string; Name?: string } };
  ReceivedDateTime?: string;
  IsRead?: boolean;
}

/** An OData string literal: single-quoted, with any `'` doubled. */
function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** An ISO 8601 argument as an OData date-time literal (UTC); a zone-less time is read as UTC. */
function dateLiteral(name: string, raw: string): string {
  const ms = parseTime(raw);
  if (ms === undefined) {
    throw new McpToolError(`\`${name}\` is not a valid ISO 8601 date or date-time: ${raw}`, {
      hint: 'Pass e.g. "2026-08-01" or "2026-08-01T00:00:00Z" (a time with no zone is read as UTC).',
    });
  }
  return utcLiteral(ms);
}

/**
 * The `$filter` for what Outlook can evaluate itself: senders (OR-joined),
 * the received-time window (`receivedAfter` inclusive, `receivedBefore`
 * exclusive) and read state, AND-joined. The subject conditions are not here
 * — they are applied to each listed row — so a subject-only filter lists the
 * whole folder (`undefined`).
 */
export function buildMoveFilter(f: MoveFilter): string | undefined {
  const parts: string[] = [];
  if (f.from?.length) {
    parts.push(`(${f.from.map((a) => `From/EmailAddress/Address eq ${odataString(a)}`).join(' or ')})`);
  }
  if (f.receivedBefore !== undefined) parts.push(`ReceivedDateTime lt ${dateLiteral('receivedBefore', f.receivedBefore)}`);
  if (f.receivedAfter !== undefined) parts.push(`ReceivedDateTime ge ${dateLiteral('receivedAfter', f.receivedAfter)}`);
  if (f.isRead !== undefined) parts.push(`IsRead eq ${f.isRead}`);
  return parts.length ? parts.join(' and ') : undefined;
}

/**
 * Whether a listed row meets the conditions Outlook did not (or may not
 * exactly) evaluate: the subject include/exclude lists, and — as a guard —
 * the sender, compared ignoring case.
 */
function rowMatches(row: Row, f: MoveFilter): boolean {
  if (f.from?.length) {
    const sender = row.From?.EmailAddress?.Address?.toLowerCase();
    if (!sender || !f.from.some((a) => a.toLowerCase() === sender)) return false;
  }
  const subject = (row.Subject ?? '').toLowerCase();
  if (f.subjectContains?.length && !f.subjectContains.some((s) => subject.includes(s.toLowerCase()))) return false;
  if (f.exceptSubjectContains?.some((s) => subject.includes(s.toLowerCase()))) return false;
  return true;
}

const clip = (s: string | undefined) =>
  s === undefined ? undefined : s.length > SUBJECT_MAX ? `${s.slice(0, SUBJECT_MAX - 1)}…` : s;

/**
 * List the folder with the server-side filter, follow every nextLink, and
 * keep the first `max` matching rows while counting all of them — so the
 * preview can say how many lie beyond `max`. Reads only.
 */
async function scan(
  client: OutlookClient,
  folder: string,
  f: MoveFilter,
  max: number,
): Promise<{ rows: (Row & { Id: string })[]; matched: number }> {
  const filter = buildMoveFilter(f);
  const query: QueryParams = { $filter: filter, $select: SELECT, $top: PAGE_SIZE };
  type Page = { value?: Row[]; '@odata.nextLink'?: string };
  let page: Page | undefined = await client.get<Page>(`/me/mailfolders/${encodeURIComponent(folder)}/messages`, {
    query,
  });
  const rows: (Row & { Id: string })[] = [];
  const seen = new Set<string>();
  let matched = 0;
  for (;;) {
    for (const row of page?.value ?? []) {
      if (typeof row.Id !== 'string' || seen.has(row.Id) || !rowMatches(row, f)) continue;
      seen.add(row.Id);
      matched++;
      if (rows.length < max) rows.push(row as Row & { Id: string });
    }
    const next: string | undefined = page?.['@odata.nextLink'];
    if (!next) break;
    throwIfCancelled();
    page = await client.getAbsolute<Page>(next);
  }
  return { rows, matched };
}

/**
 * A folder's DisplayName for the preview and the result, so the user sees
 * "Alerts" rather than a raw id. Well-known names ("inbox", "archive")
 * resolve the same way. Any failure falls back to what was passed — the name
 * is only for display, never for the move itself.
 */
async function folderName(client: OutlookClient, id: string): Promise<string> {
  try {
    const folder = await client.get<{ DisplayName?: string }>(`/me/mailfolders/${encodeURIComponent(id)}`, {
      query: { $select: 'DisplayName' },
    });
    const name = folder?.DisplayName;
    return typeof name === 'string' && name.trim() ? name : id;
  } catch (e) {
    if (isCredentialFailure(e)) throw e;
    return id;
  }
}

/** A short, stable digest of the ids that will move, bound into the token. */
function idsDigest(ids: string[]): string {
  return createHash('sha256').update(ids.join('\n')).digest('base64url').slice(0, 22);
}

type Outcome =
  | { kind: 'moved'; markedRead?: boolean; markReadFailed?: boolean }
  | { kind: 'failed'; id: string; subject?: string; error: string }
  | { kind: 'unknown'; id: string; error: WriteOutcomeUnknownError }
  | { kind: 'skipped' };

const DESCRIPTION =
  'Move every message in a folder (default the inbox) that matches a filter to another folder — e.g. "archive", "deleteditems", or a folder id from outlook_list_folders — under ONE confirmation for the whole set. ' +
  'Reversible: it only moves, never deletes. ' +
  'Filter by any of `from` (sender addresses, exact, ignoring case), `receivedBefore` / `receivedAfter` (ISO 8601; before is exclusive, after inclusive), `isRead`, and `subjectContains` (any of); at least one is required, so an unfiltered "move everything" is refused. `exceptSubjectContains` (none of) narrows further. ' +
  'To run an inbox rule over mail that arrived before it existed, pass the rule\'s sender list as `from`, its exceptions as `exceptSubjectContains`, its folder as `destination`, and `markRead: true` if the rule marks mail read. ' +
  `It moves at most \`max\` messages per call (default ${BULK_MOVE_DEFAULT_MAX}, at most ${BULK_MOVE_HARD_MAX}); when more match, the preview says so, the result reports \`remaining\`, and calling again moves the next set. ` +
  'The preview shows the number matched, the filter, the destination and a sample of 10. Every call lists the folder afresh, so if the matching mail changed before approval the result is a new preview to approve rather than a move. ' +
  'One failed message does not stop the rest: the result reports `moved`, `failedCount` with the first 20 failures, `markedRead`, `remaining` and `unknown`. Moved messages get NEW ids. ' +
  CONFIRM_FLOW_SENTENCE +
  ' ' +
  TIMEOUT_SENTENCE.replace('the result has status "unknown"', 'that message is counted under `unknown`') +
  ' ' +
  CONFIRM_INJECTION_RULE;

export function registerBulkTools(server: McpServer, client: OutlookClient): void {
  server.registerTool(
    'outlook_move_messages',
    {
      description: DESCRIPTION,
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      _meta: workMeta('change', {
        aliases: ['bulk move', 'archive old mail', 'apply rule', 'run rules now', 'clean up inbox'],
      }),
      inputSchema: z.object({
        folder: z
          .string()
          .min(1)
          .optional()
          .describe('Source folder id or well-known name (default "inbox")'),
        destination: z.string().min(1).describe('Destination folder id or well-known name (e.g. "archive")'),
        from: z
          .array(z.string().min(3))
          .min(1)
          .optional()
          .describe('Sender addresses; a message from any of them matches (exact address, ignoring case)'),
        receivedBefore: z
          .string()
          .min(1)
          .optional()
          .describe('Only mail received before this ISO 8601 date/time (no zone = UTC), e.g. "2026-08-14"'),
        receivedAfter: z
          .string()
          .min(1)
          .optional()
          .describe('Only mail received at or after this ISO 8601 date/time (no zone = UTC)'),
        isRead: z.boolean().optional().describe('true: only read mail; false: only unread mail'),
        subjectContains: z
          .array(z.string().min(1))
          .min(1)
          .optional()
          .describe('Subject contains ANY of these (ignoring case)'),
        exceptSubjectContains: z
          .array(z.string().min(1))
          .min(1)
          .optional()
          .describe('Skip mail whose subject contains ANY of these (ignoring case)'),
        markRead: z.boolean().optional().describe('Also mark each moved message read (as a rule that marks read would)'),
        max: z
          .number()
          .int()
          .min(1)
          .max(BULK_MOVE_HARD_MAX)
          .optional()
          .describe(`Most messages to move in this call (default ${BULK_MOVE_DEFAULT_MAX}, at most ${BULK_MOVE_HARD_MAX})`),
        confirmToken: confirmTokenParam,
      }),
    },
    async (args, ctx) => {
      const { destination, markRead, confirmToken } = args;
      const folder = args.folder ?? 'inbox';
      const max = args.max ?? BULK_MOVE_DEFAULT_MAX;
      const filter: MoveFilter = {
        from: args.from,
        receivedBefore: args.receivedBefore,
        receivedAfter: args.receivedAfter,
        isRead: args.isRead,
        subjectContains: args.subjectContains,
        exceptSubjectContains: args.exceptSubjectContains,
      };
      if (
        !filter.from?.length &&
        filter.receivedBefore === undefined &&
        filter.receivedAfter === undefined &&
        filter.isRead === undefined &&
        !filter.subjectContains?.length
      ) {
        throw new McpToolError('Refusing to move every message in the folder: give at least one condition.', {
          hint: 'Pass at least one of from, receivedBefore, receivedAfter, isRead or subjectContains.',
        });
      }
      // Validates the dates before anything is read.
      buildMoveFilter(filter);

      const { rows, matched } = await scan(client, folder, filter, max);
      if (matched === 0) {
        return minifiedResult({
          moved: 0,
          matched: 0,
          destination,
          note: 'Nothing in the folder matches; nothing was moved.',
        });
      }

      const [sourceName, destinationName] = await Promise.all([
        folderName(client, folder),
        folderName(client, destination),
      ]);
      const ids = rows.map((r) => r.Id);
      const beyondMax = matched - rows.length;
      const shownFilter = { folder, ...filter };
      const gate = await confirmWrite(ctx, {
        tool: 'outlook_move_messages',
        action: 'mail.bulk_move',
        message: 'Review and confirm this bulk move:',
        summary: `Move ${rows.length} message${rows.length === 1 ? '' : 's'} from ${sourceName} to ${destinationName}${markRead ? ' and mark them read' : ''}`,
        // One signed-in mailbox per server process.
        account: undefined,
        target: folder,
        // The set that will move: a change between preview and approval is a new preview.
        revision: `${matched}:${idsDigest(ids)}`,
        request: {
          method: 'POST',
          path: '/me/messages/{id}/move',
          body: { DestinationId: destination },
        },
        preview: {
          matched,
          willMove: rows.length,
          ...(beyondMax > 0
            ? {
                beyondMax,
                note: `${matched} messages match, more than ${max}: only the first ${max} move now. Call again afterwards to move the rest.`,
              }
            : {}),
          filter: shownFilter,
          source: sourceName,
          sourceId: folder,
          destination: destinationName,
          destinationId: destination,
          markRead: markRead === true,
          sample: rows.slice(0, SAMPLE_SIZE).map((r) => ({
            from: r.From?.EmailAddress?.Address,
            subject: clip(r.Subject),
            received: r.ReceivedDateTime,
          })),
        },
        args: { folder, destination, ...filter, markRead: markRead === true, max },
        confirmToken,
      });
      if (gate) return gate;

      // Past the gate: never stop the whole run for one message. Only the
      // caller cancelling, or a credential failing for everyone, stops new moves.
      let halted: string | undefined;
      let done = 0;
      const outcomes = await mapWithConcurrency(rows, MOVE_CONCURRENCY, async (row): Promise<Outcome> => {
        if (halted !== undefined) return { kind: 'skipped' };
        if (currentCallSignal()?.aborted) {
          halted = 'cancelled';
          return { kind: 'skipped' };
        }
        let newId: string | undefined;
        try {
          const moved = await client.write<{ Id?: string }>(
            'POST',
            `/me/messages/${encodeURIComponent(row.Id)}/move`,
            { DestinationId: destination },
          );
          newId = moved?.Id;
        } catch (e) {
          // Sent, but no answer: it may well have moved. Never retried.
          if (e instanceof WriteOutcomeUnknownError) return { kind: 'unknown', id: row.Id, error: e };
          if (isCredentialFailure(e)) halted = e instanceof Error ? e.message : String(e);
          return { kind: 'failed', id: row.Id, subject: clip(row.Subject), error: e instanceof Error ? e.message : String(e) };
        } finally {
          done++;
          if (done % 25 === 0) await reportProgress(done, rows.length, `${done} of ${rows.length} moved`);
        }
        if (!markRead || row.IsRead === true) return { kind: 'moved' };
        // Moved, but Outlook returned no new id to mark read: say so rather than count it as done.
        if (!newId) return { kind: 'moved', markReadFailed: true };
        try {
          await client.write('PATCH', `/me/messages/${encodeURIComponent(newId)}`, { IsRead: true });
          return { kind: 'moved', markedRead: true };
        } catch {
          return { kind: 'moved', markReadFailed: true };
        }
      });

      const moved = outcomes.filter((o) => o.kind === 'moved');
      const failed = outcomes.filter((o): o is Extract<Outcome, { kind: 'failed' }> => o.kind === 'failed');
      const unknown = outcomes.filter((o): o is Extract<Outcome, { kind: 'unknown' }> => o.kind === 'unknown');
      const skipped = outcomes.filter((o) => o.kind === 'skipped').length;
      const markReadFailed = moved.filter((o) => o.markReadFailed).length;
      const result = {
        moved: moved.length,
        failed: failed.slice(0, LISTED_FAILURES).map(({ id, subject, error }) => ({ id, subject, error })),
        failedCount: failed.length,
        ...(markRead ? { markedRead: moved.filter((o) => o.markedRead).length } : {}),
        ...(markReadFailed ? { markReadFailed } : {}),
        ...(unknown.length
          ? {
              unknown: {
                count: unknown.length,
                ids: unknown.slice(0, LISTED_FAILURES).map((u) => u.id),
                warning: unknownOutcome(unknown[0]!.error, `the ${destinationName} folder (outlook_list_messages)`).warning,
              },
            }
          : {}),
        remaining: beyondMax,
        destination: destinationName,
        destinationId: destination,
        ...(skipped
          ? {
              notAttempted: skipped,
              stoppedBecause: halted === 'cancelled' ? 'the call was cancelled' : `a credential failure: ${halted}`,
            }
          : {}),
        note: 'Moved messages have new ids; the old ones no longer resolve.',
      };
      const out = minifiedResult(result);
      return moved.length === 0 && failed.length > 0 ? { ...out, isError: true } : out;
    },
  );
}
