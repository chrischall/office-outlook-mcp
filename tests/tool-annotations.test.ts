import { describe, expect, it } from 'vitest';
import { registerMailTools } from '../src/tools/mail.js';
import { registerCalendarTools } from '../src/tools/calendar.js';
import { registerDirectoryTools } from '../src/tools/directory.js';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';
import type { OutlookClient } from '../src/client.js';

/**
 * The fleet annotation meta-test, read off the REGISTERED config rather than a
 * hand-kept list. `destructiveHint` defaults to TRUE whenever readOnlyHint is
 * false, so a write that forgets it is published as destructive and nothing
 * fails; and a tool with no `openWorldHint` leaves a client to guess whether it
 * reaches the network. Every tool here talks to Outlook, so all are open-world.
 */
interface Ann {
  readOnlyHint?: unknown;
  destructiveHint?: unknown;
  idempotentHint?: unknown;
  openWorldHint?: unknown;
}

function registeredAnnotations(): Record<string, Ann | undefined> {
  const seen: Record<string, Ann | undefined> = {};
  const server = {
    registerTool: (name: string, cfg: { annotations?: Ann }) => {
      seen[name] = cfg.annotations;
    },
  } as never;
  const client = {} as OutlookClient;
  registerMailTools(server, client);
  registerCalendarTools(server, client);
  registerDirectoryTools(server, client);
  registerWriteTools(server, client);
  registerHealthcheckTool(server, client);
  return seen;
}

describe('tool annotations', () => {
  it('covers the full surface (guards against a registrar being dropped here)', () => {
    expect(Object.keys(registeredAnnotations())).toHaveLength(22);
  });

  it('sets an explicit boolean readOnlyHint and openWorldHint on every tool', () => {
    const missing = Object.entries(registeredAnnotations())
      .filter(([, a]) => typeof a?.readOnlyHint !== 'boolean' || typeof a?.openWorldHint !== 'boolean')
      .map(([name]) => name);
    expect(missing).toEqual([]);
  });

  it('marks every tool open-world: each one calls the Outlook API', () => {
    const closed = Object.entries(registeredAnnotations())
      .filter(([, a]) => a?.openWorldHint !== true)
      .map(([name]) => name);
    expect(closed).toEqual([]);
  });

  it('sets an explicit boolean destructiveHint on every write', () => {
    const undeclared = Object.entries(registeredAnnotations())
      .filter(([, a]) => a?.readOnlyHint === false && typeof a?.destructiveHint !== 'boolean')
      .map(([name]) => name);
    expect(undeclared).toEqual([]);
  });

  it('never lets a read claim to be destructive', () => {
    const contradictory = Object.entries(registeredAnnotations())
      .filter(([, a]) => a?.readOnlyHint === true && a?.destructiveHint === true)
      .map(([name]) => name);
    expect(contradictory).toEqual([]);
  });

  it('classifies each write by the inverse test', () => {
    const ann = registeredAnnotations();
    const destructive = (name: string) => ann[name]?.destructiveHint;
    // Reach another person: mail, invitations and updated invitations cannot be unsent.
    expect(destructive('outlook_send_mail')).toBe(true);
    expect(destructive('outlook_create_event')).toBe(true);
    expect(destructive('outlook_update_event')).toBe(true);
    // No tool in this set deletes a draft, so nothing restores the prior state.
    expect(destructive('outlook_create_draft')).toBe(true);
    // Self-inverse: mark_read(isRead: !x) and move_message back to the old folder.
    expect(destructive('outlook_mark_read')).toBe(false);
    expect(destructive('outlook_move_message')).toBe(false);
  });
});
