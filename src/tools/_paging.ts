import { z } from 'zod';
import { McpToolError } from '@chrischall/mcp-utils';
import type { OutlookClient, QueryParams } from '../client.js';
import { stripOData } from '../view.js';

/**
 * Paging for every list tool.
 *
 * Outlook caps each page (`$top`) and signals the rest with
 * `@odata.nextLink`. A listing that drops that link is indistinguishable from a
 * complete one, so the model says with confidence that an item does not
 * exist. Every list tool therefore returns `nextLink` when there is more, and
 * accepts it back to fetch the next page.
 */
export const nextLinkParam = z
  .string()
  .min(1)
  .optional()
  .describe(
    'Continue a previous listing: pass its `nextLink` value unchanged to fetch the next page. When set, the paging/filter arguments are taken from the link and ignored here.',
  );

/**
 * Fetch the first page at `path` with `query`, or the page a previous
 * `nextLink` points at. A link already carries its own query, so `query` is
 * dropped when following one.
 */
export function fetchPage<T>(
  client: OutlookClient,
  nextLink: string | undefined,
  path: string,
  opts: { text?: boolean; prefer?: string; query?: QueryParams } = {},
): Promise<T> {
  if (nextLink === undefined) return client.get<T>(path, opts);
  const { query: _ignored, ...headerOpts } = opts;
  return client.getAbsolute<T>(nextLink, headerOpts);
}

/**
 * Pass a collection through without a typed projection, trimming the OData
 * envelope from each item but keeping the paging link.
 */
export function plainCollection(data: {
  value?: Record<string, unknown>[];
  '@odata.nextLink'?: string;
}): Record<string, unknown> {
  const items = Array.isArray(data?.value) ? data.value : [];
  const out: Record<string, unknown> = { count: items.length, items: items.map(stripOData) };
  const next = data?.['@odata.nextLink'];
  if (next) out.nextLink = next;
  return out;
}

/**
 * Oldest-first keyset paging from a UTC time (`from` in, `next_from` out).
 *
 * A position is the newest `ReceivedDateTime` already returned plus the Ids
 * returned AT that instant, so a page boundary that splits messages sharing a
 * timestamp neither skips nor repeats one. Keyset rather than `$skip` on
 * purpose: a caller that marks mail read between get_unread pages removes rows
 * from the filtered set, which shifts every later `$skip` offset and silently
 * loses messages. A time bound does not move.
 */
export interface FromPosition {
  /** ISO 8601 UTC time of the newest message returned so far. */
  t: string;
  /** Ids of the returned messages whose ReceivedDateTime equals `t`. */
  ids: string[];
}

export const fromParam = z
  .string()
  .min(1)
  .optional()
  .describe(
    'Page oldest first from this point: an ISO 8601 UTC time (e.g. "2026-10-01T00:00:00Z"; one with no zone is read as UTC), or a previous result\'s `next_from` passed unchanged. The result always carries `next_from` and `has_more`.',
  );

/** Outlook's per-request `$top` ceiling. */
const MAX_TOP = 1000;

const FROM_HINT =
  'Pass a previous next_from unchanged, or an ISO 8601 time such as 2026-10-01T00:00:00Z (one with no zone is read as UTC).';

/** An ISO 8601 date, or date-time with or without a zone. */
const ISO_DATE_TIME = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/i;

/**
 * The OData literal for `t` (unquoted), and the canonical form kept in a
 * token. Milliseconds are dropped only when zero — truncating a real fraction
 * would widen `ge` to messages older than the position and repeat them.
 */
function utcLiteral(ms: number): string {
  return new Date(ms).toISOString().replace(/\.000Z$/, 'Z');
}

function parseTime(raw: string): number | undefined {
  if (!ISO_DATE_TIME.test(raw)) return undefined;
  // A date-time with no zone would parse as LOCAL time; `from` is UTC, so it is pinned there.
  // (A bare date already parses as UTC.)
  const t = /T[\d:.]+$/.test(raw) ? `${raw}Z` : raw;
  const ms = Date.parse(t);
  if (Number.isNaN(ms)) return undefined;
  // Date.parse rolls 2026-02-30 over to March; a round trip of the date part catches it.
  if (new Date(ms + offsetMs(t)).toISOString().slice(0, 10) !== t.slice(0, 10)) return undefined;
  return ms;
}

/** The zone offset written in `t`, so the calendar date can be checked as written. */
function offsetMs(t: string): number {
  const m = /([+-])(\d{2}):?(\d{2})$/.exec(t);
  if (!m || /Z$/i.test(t)) return 0;
  const mins = Number(m[2]) * 60 + Number(m[3]);
  return (m[1] === '-' ? -1 : 1) * mins * 60_000;
}

export function encodeFrom(pos: FromPosition): string {
  return Buffer.from(JSON.stringify({ t: pos.t, ids: pos.ids }), 'utf8').toString('base64url');
}

/** Read `from`: an ISO time, or a `next_from` token. Anything else is a tool error with a hint. */
export function decodeFrom(from: string): FromPosition {
  if (/^\d{4}-/.test(from)) {
    const ms = parseTime(from);
    if (ms === undefined) {
      throw new McpToolError(`\`from\` is not a valid ISO 8601 UTC date-time: ${from}`, { hint: FROM_HINT });
    }
    return { t: utcLiteral(ms), ids: [] };
  }
  let parsed: unknown;
  try {
    parsed = /^[A-Za-z0-9_-]+$/.test(from) ? JSON.parse(Buffer.from(from, 'base64url').toString('utf8')) : undefined;
  } catch {
    parsed = undefined;
  }
  const p = parsed as { t?: unknown; ids?: unknown } | undefined;
  const ms = typeof p?.t === 'string' ? parseTime(p.t) : undefined;
  if (
    ms === undefined ||
    !Array.isArray(p?.ids) ||
    !p.ids.every((id): id is string => typeof id === 'string')
  ) {
    throw new McpToolError('`from` is neither an ISO 8601 time nor a next_from token.', { hint: FROM_HINT });
  }
  return { t: utcLiteral(ms), ids: p.ids };
}

/**
 * The `$filter` date bound and `$top` for a page from `pos`. The bound is
 * inclusive (`ge`), so the tied ids come back again: ask for that many extra
 * rows so dropping them still leaves a full page.
 */
export function fromQuery(pos: FromPosition, limit: number): { bound: string; top: number } {
  return { bound: `ReceivedDateTime ge ${pos.t}`, top: Math.min(limit + pos.ids.length, MAX_TOP) };
}

/**
 * Cut an ascending page fetched with `fromQuery` down to what the caller
 * sees: drop the rows already returned at `pos.t`, keep the first `limit`,
 * and work out the next position. `has_more` is true when rows were left over
 * or the API has another page.
 */
export function fromPage<T extends { Id?: string; ReceivedDateTime?: string }>(
  rows: T[],
  pos: FromPosition,
  limit: number,
  apiHasMore: boolean,
): { rows: T[]; next_from: string; has_more: boolean } {
  const seen = new Set(pos.ids);
  const fresh = rows.filter((r) => r.Id === undefined || !seen.has(r.Id));
  const kept = fresh.slice(0, limit);
  const next = nextPosition(kept, pos);
  return { rows: kept, next_from: encodeFrom(next), has_more: fresh.length > limit || apiHasMore };
}

function nextPosition(rows: { Id?: string; ReceivedDateTime?: string }[], pos: FromPosition): FromPosition {
  const dated = rows
    .map((r) => ({ id: r.Id, ms: r.ReceivedDateTime ? Date.parse(r.ReceivedDateTime) : NaN }))
    .filter((r): r is { id: string; ms: number } => r.id !== undefined && !Number.isNaN(r.ms));
  if (dated.length === 0) return pos;
  const newest = Math.max(...dated.map((r) => r.ms));
  const tied = dated.filter((r) => r.ms === newest).map((r) => r.id);
  // The position did not move: earlier ties at this instant stay excluded.
  const carried = newest === Date.parse(pos.t) ? pos.ids.filter((id) => !tied.includes(id)) : [];
  return { t: utcLiteral(newest), ids: [...carried, ...tied] };
}
