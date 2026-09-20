/**
 * Global test setup.
 *
 * Pins the token cache into a throwaway directory and turns it off, so no test
 * can read or write the developer's real `~/.office-outlook-mcp` — or, worse,
 * trigger a live browser capture. Every credential in the suite is injected.
 */
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const dir = mkdtempSync(join(tmpdir(), 'office-outlook-mcp-test-'));

process.env.OUTLOOK_TOKEN_FILE = join(dir, 'token.json');
process.env.OUTLOOK_TOKEN_CACHE = 'false';
// Guard against a test accidentally reaching the bridge: without a declared
// token the client should surface a config error, never open a socket.
delete process.env.OUTLOOK_ACCESS_TOKEN;
