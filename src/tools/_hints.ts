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
  '"brief": one line per item, `title — <ISO date or "no date"> — <status>`, at most 120 chars each.';

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

/** `title — date — status`, the title cut with an ellipsis to keep the line ≤ 120. */
export function briefLine(title: string | undefined, date: string, status: string): string {
  const tail = ` — ${date} — ${status}`;
  const shown = flat(title) ?? '(no subject)';
  const room = BRIEF_MAX - tail.length;
  return `${shown.length > room ? `${shown.slice(0, Math.max(room - 1, 0))}…` : shown}${tail}`;
}

/** One line per item, as plain text — never a JSON envelope. */
export function briefResult(lines: string[]): CallToolResult {
  return { content: [{ type: 'text', text: lines.join('\n') }] };
}

/** `NotStarted` → `not started`, `TentativelyAccepted` → `tentatively accepted`. */
export function spaced(value: string): string {
  return value.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase();
}
