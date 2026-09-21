import { describe, expect, it } from 'vitest';
import { spawn } from 'node:child_process';
import { mkdtempSync, copyFileSync, writeFileSync, existsSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const bundle = join(root, 'dist', 'bundle.js');
const binEntry = join(root, 'dist', 'index.js');

/** Drive the stdio handshake against a built entrypoint and return its tools. */
async function handshake(entry: string, cwd: string): Promise<string[]> {
  const child = spawn(process.execPath, [entry], {
    cwd,
    stdio: ['pipe', 'pipe', 'pipe'],
    // No credential: the server must still boot and answer tools/list, which is
    // what a host's install-time probe does.
    env: { ...process.env, OUTLOOK_DISABLE_FETCHPROXY: '1' },
  });

  const out: string[] = [];
  const err: string[] = [];
  child.stdout.setEncoding('utf8');
  child.stderr.setEncoding('utf8');
  child.stdout.on('data', (d: string) => out.push(d));
  child.stderr.on('data', (d: string) => err.push(d));

  for (const msg of [
    { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } } },
    { jsonrpc: '2.0', method: 'notifications/initialized' },
    { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} },
  ]) {
    child.stdin.write(`${JSON.stringify(msg)}\n`);
  }

  const names = await new Promise<string[]>((resolve, reject) => {
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`timed out; stderr: ${err.join('').slice(0, 400)}`));
    }, 30_000);
    const tryParse = () => {
      for (const line of out.join('').split('\n')) {
        if (!line.trim()) continue;
        try {
          const m = JSON.parse(line) as { id?: number; result?: { tools?: { name: string }[] } };
          if (m.id === 2 && m.result?.tools) {
            clearTimeout(timer);
            child.kill('SIGTERM');
            resolve(m.result.tools.map((t) => t.name).sort());
            return;
          }
        } catch {
          /* partial line; wait for more */
        }
      }
    };
    child.stdout.on('data', tryParse);
    child.on('exit', () => {
      clearTimeout(timer);
      reject(new Error(`exited early; stderr: ${err.join('').slice(0, 400)}`));
    });
  });
  return names;
}

describe('built server boots', () => {
  it.runIf(existsSync(bundle))(
    'loads the bundle with NO node_modules — the .mcpb runtime',
    async () => {
      // The regression this exists for: an eager import of an esbuild-external
      // dependency throws ERR_MODULE_NOT_FOUND before the server can answer
      // initialize, and the host reports only "transport closed unexpectedly".
      const dir = mkdtempSync(join(tmpdir(), 'office-outlook-mcpb-'));
      copyFileSync(bundle, join(dir, 'bundle.js'));
      writeFileSync(join(dir, 'package.json'), '{"type":"module"}');
      const tools = await handshake(join(dir, 'bundle.js'), dir);
      // A floor, not an exact count: PR CI runs the branch merged with main, so
      // a hardcoded length breaks the moment another PR adds a tool.
      expect(tools.length).toBeGreaterThanOrEqual(18);
      expect(tools).toContain('outlook_healthcheck');
      expect(tools).toContain('outlook_list_messages');
    },
    60_000,
  );

  it.runIf(existsSync(binEntry))(
    'loads the package bin entrypoint with node_modules present',
    async () => {
      // Catches a wrong `rootDir` emitting dist/src/index.js while `bin` points
      // at dist/index.js.
      const tools = await handshake(binEntry, root);
      expect(tools.length).toBeGreaterThanOrEqual(18);
    },
    60_000,
  );

  it.runIf(existsSync(bundle))(
    'boots from the repo .mcp.json the way a project-scoped config does',
    async () => {
      // The regression this exists for: `.mcp.json` launched
      // `${CLAUDE_PLUGIN_ROOT}/dist/bundle.js`, which a PLUGIN install defines
      // and a project-scoped load does NOT. Claude Code ran `node
      // /dist/bundle.js`, the process died instantly, and the session reported
      // only "CONNECTION_CLOSED" — auth looked broken when nothing had started.
      // The plugin's own config keeps the variable; this file must not need it.
      const cfg = JSON.parse(readFileSync(join(root, '.mcp.json'), 'utf8')) as {
        mcpServers: Record<string, { command: string; args: string[] }>;
      };
      const server = cfg.mcpServers.outlook;
      expect(JSON.stringify(server)).not.toContain('CLAUDE_PLUGIN_ROOT');
      // Resolved against the repo root, exactly as a project-scoped launch does.
      const entry = join(root, server.args[server.args.length - 1]);
      expect(existsSync(entry)).toBe(true);
      const tools = await handshake(entry, root);
      expect(tools).toContain('outlook_healthcheck');
    },
    60_000,
  );
});
