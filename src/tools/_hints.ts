/**
 * Context-group hints, published on each tool as `_meta.mint` (mcp-host's
 * `mintHintSchema`, spec 2026-11-03 D4), and the `view: "brief"` rung the
 * canonical reads answer a group's question in. Inert data to any other
 * client: a host that does not read them sees an extra `_meta` key.
 *
 * The shape is fixed by the host's contract, so it is restated here rather
 * than imported — this server takes no dependency on mcp-host. `when` names
 * THIS tool's own argument names (the host fills them), never values;
 * `aliases` are the words a person says, lowercase, at most 8, each at most
 * 40 chars. No tool here names a `subject`: one signed-in mailbox has no
 * listable set of subjects to fan out over.
 */
import { z } from 'zod';
import type { CallToolResult } from '@modelcontextprotocol/server';
import { viewParam, type View } from '@chrischall/mcp-utils';
import { mailboxUntrusted } from './_untrusted.js';

export type MintIntent =
  | 'whats-new' | 'whats-due' | 'whats-on' | 'lookup' | 'turn-on-off' | 'send' | 'change' | 'other';

export interface MintHint {
  domain: 'work';
  intent: MintIntent;
  canonical?: true;
  subject?: string;
  when?: { start: string; end: string };
  aliases?: string[];
}

/** `_meta` for an Outlook tool: every tool here is in the `work` group. */
export function workMeta(
  intent: MintIntent,
  extra: Omit<MintHint, 'domain' | 'intent'> = {},
): { mint: MintHint } {
  return { mint: { domain: 'work', intent, ...extra } };
}

/** The longest a brief line may be. */
export const BRIEF_MAX = 120;

const BRIEF_NOTE =
  '"brief": one line per item, `title — <ISO date or "no date"> — <status>`, at most 120 chars each, as `data` in the untrusted-content envelope; a line starting `-- ` is a note (`-- none` for an empty page; a title starting `-- ` or `\\` gets one leading `\\`); a last line `-- more` means another page exists (re-ask without "brief" for its nextLink).';

/**
 * The `view` parameter of a canonical read: the tool's own rungs plus
 * `brief`. Absent still means the tool's default.
 */
export function briefViewParam(rungs: readonly View[], note: string) {
  const base = viewParam(rungs);
  return z
    .enum([...rungs, 'brief'] as readonly string[] as [string, ...string[]])
    .optional()
    .describe(`${base.description ?? ''} ${BRIEF_NOTE} ${note}`.trim());
}

/** Whitespace in a third party's text collapsed, so one item stays one line. */
function flat(s: string | undefined): string | undefined {
  const out = s?.replace(/\s+/g, ' ').trim();
  return out ? out : undefined;
}

/** Outlook pads times to 7 fractional digits; that is noise in one line. */
export function briefDate(dt: string | undefined): string {
  return flat(dt?.replace(/\.0+(?=Z?$)/, '')) ?? 'no date';
}

/** The shortest a title is cut to before the status gives way too. */
const TITLE_FLOOR = 24;

const cut = (s: string, n: number): string => (s.length > n ? `${s.slice(0, Math.max(n - 1, 0))}…` : s);

/**
 * `title — date — status`, cut to ≤ 120 chars as a WHOLE, title first: the
 * title shortens until the line fits; once it is at its floor the status is
 * cut too (and, past that, the line itself). A title that starts with `-- `
 * or `\` is escaped with one leading `\`, so a subject or sender name can
 * never read as a note — mcp-host strips that one backslash from every item
 * line after its note check.
 */
export function briefLine(title: string | undefined, date: string, status: string): string {
  const raw = flat(title) ?? '(no subject)';
  const shown = raw.startsWith('-- ') || raw.startsWith('\\') ? `\\${raw}` : raw;
  const sep = ` — ${date} — `;
  const room = BRIEF_MAX - sep.length;
  if (shown.length + status.length <= room) return `${shown}${sep}${status}`;
  const titleRoom = Math.max(room - status.length, Math.min(shown.length, TITLE_FLOOR));
  const t = cut(shown, titleRoom);
  return cut(`${t}${sep}${cut(status, room - t.length)}`, BRIEF_MAX);
}

/** The note line a brief answer ends with when Outlook holds another page. */
export const BRIEF_MORE = '-- more';

/** The single note line a brief answer with no items carries. */
export const BRIEF_NONE = '-- none';

/**
 * One line per item, inside the untrusted-data envelope as `data: string[]`
 * (mcp-host's brief-view contract): sender names, subjects and event text are
 * written by other people, so a brief line is fenced like every other rung of
 * these tools. A line starting with `-- ` is a NOTE, not an item: `-- none`
 * answers an empty page, and `-- more`
 * says Outlook has another page (it gives no total, so no `N of M`) — the
 * brief view never silently drops one. Re-ask without `view: "brief"` for the
 * `nextLink` / `next_from` to follow it.
 */
export function briefResult(lines: string[], hasMore = false): CallToolResult {
  const items = lines.length > 0 ? lines : [BRIEF_NONE];
  return mailboxUntrusted(hasMore ? [...items, BRIEF_MORE] : items);
}

/** `NotStarted` → `not started`, `TentativelyAccepted` → `tentatively accepted`. */
export function spaced(value: string): string {
  return value.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase();
}
