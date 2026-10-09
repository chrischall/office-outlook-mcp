/**
 * The mailbox's own Windows time-zone name.
 *
 * Outlook's calendar endpoints default to UTC, which is wrong twice over: a
 * read shows a 7:15am Eastern meeting as 11:15 behind nothing louder than
 * `TimeZone: "UTC"`, and a WRITE books "3pm" four hours off in a real
 * calendar. The mailbox already declares its zone, so the default is knowable
 * rather than guessable — both sides ask here.
 *
 * Cached because it is a per-mailbox constant and paying a round trip per call
 * to re-learn it is not worth it. Keyed by CLIENT rather than held in a module
 * variable: the zone belongs to the mailbox, not the process, so a global
 * would hand one mailbox's zone to another client in the same process — and
 * would leak between tests in file order, which is how a test that believes it
 * exercised the failure path quietly stops doing so.
 *
 * For a READ a failure is swallowed deliberately — the zone is a default, not
 * a requirement, and must not take the calendar down with it. A WRITE passes
 * `required: true` and gets an error instead: there the fallback is booking
 * the caller's "3pm" in UTC, with invitations already sent. Either way a
 * failure is NOT cached, so a transient blip does not pin the mailbox to UTC
 * for the life of the server.
 */
import { McpToolError } from '@chrischall/mcp-utils';
import type { OutlookClient } from './client.js';

const zoneByClient = new WeakMap<OutlookClient, string>();

export async function mailboxTimeZone(
  client: OutlookClient,
  opts: { required?: boolean } = {},
): Promise<string | undefined> {
  const cached = zoneByClient.get(client);
  if (cached) return cached;
  try {
    const settings = await client.get<{ TimeZone?: string }>('/me/MailboxSettings');
    const zone = settings?.TimeZone?.trim() || undefined;
    if (zone) zoneByClient.set(client, zone);
    return zone;
  } catch (e) {
    if (!opts.required) return undefined;
    throw new McpToolError("Could not read the mailbox's time zone, so the event time is ambiguous.", {
      hint:
        'Retry, or pass `timeZone` explicitly (a Windows zone name such as "Eastern Standard Time"). ' +
        `Cause: ${e instanceof Error ? e.message : String(e)}`,
    });
  }
}
