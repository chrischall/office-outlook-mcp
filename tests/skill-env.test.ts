import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, statSync, writeFileSync, chmodSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';

const ENV_SH = resolve(__dirname, '../skills/outlook-fpx/references/outlook-env.sh');

describe('skills/outlook-fpx/references/outlook-env.sh', () => {
  let dir: string;
  afterEach(() => {
    if (dir) rmSync(dir, { recursive: true, force: true });
  });

  it("outlook_token_refresh leaves the sourcing shell's umask alone", () => {
    // The file is SOURCED into the user's interactive shell. A bare `umask 077`
    // inside the function outlives it, so every file the user creates for the
    // rest of the session comes out 0600 — the token file needs it, nothing
    // else does.
    dir = mkdtempSync(join(tmpdir(), 'outlook-env-'));
    const bin = join(dir, 'bin');
    mkdirSync(bin);
    // A stand-in for the browser bridge: prints a captured header map.
    writeFileSync(
      join(bin, 'fpx'),
      '#!/bin/sh\nprintf \'{"authorization@outlook.cloud.microsoft":"Bearer test-token"}\'\n',
    );
    chmodSync(join(bin, 'fpx'), 0o755);

    const script = [
      'umask 022',
      `. "${ENV_SH}"`,
      // Never reload the developer's real Chrome tab, and skip the 3s wait.
      '_outlook_poke_tab() { :; }',
      'sleep() { :; }',
      'outlook_token_refresh 2>/dev/null || exit 9',
      'umask',
    ].join('\n');
    const out = execFileSync('bash', ['-c', script], {
      env: { ...process.env, PATH: `${bin}:${process.env.PATH}`, OUTLOOK_FPX_DIR: join(dir, 'fpx') },
      encoding: 'utf8',
    });
    expect(out.trim()).toBe('0022');
    // ...while the token file itself is still private.
    expect(statSync(join(dir, 'fpx', 'curlrc')).mode & 0o777).toBe(0o600);
  });
});
