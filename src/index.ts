#!/usr/bin/env node
import { runMcp, loadDotenvSafely } from '@chrischall/mcp-utils';
import { VERSION } from './version.js';
import { OutlookClient } from './client.js';
import { registerMailTools } from './tools/mail.js';
import { registerCalendarTools } from './tools/calendar.js';
import { registerDirectoryTools } from './tools/directory.js';
import { registerWriteTools } from './tools/writes.js';
import { registerHealthcheckTool } from './tools/healthcheck.js';

loadDotenvSafely();

/**
 * Built HERE, in the caller, not inside a registrar.
 *
 * `runMcp` constructs a server instance per served connection — and one more
 * for a `server/discover` probe that the client then discards — so the
 * registrars run more than once. Anything built inside one would be rebuilt,
 * which for this client would mean a second TokenManager and a second browser
 * capture. Registrars register; state lives in `deps`.
 *
 * The constructor also never throws on missing config, so the server boots and
 * answers an install-time `tools/list` with no credential present.
 */
const client = new OutlookClient();

await runMcp({
  name: 'office-outlook-mcp',
  version: VERSION,
  banner:
    '[office-outlook-mcp] This project was developed and is maintained by AI. Use at your own discretion.',
  deps: client,
  tools: [
    registerMailTools,
    registerCalendarTools,
    registerDirectoryTools,
    registerWriteTools,
    registerHealthcheckTool,
  ],
});
