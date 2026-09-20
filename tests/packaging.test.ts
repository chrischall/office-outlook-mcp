import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { versionSyncTest } from '@chrischall/mcp-utils/test';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerMailTools } from '../src/tools/mail.js';
import { registerCalendarTools } from '../src/tools/calendar.js';
import { registerDirectoryTools } from '../src/tools/directory.js';
import { registerWriteTools } from '../src/tools/writes.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';
import type { OutlookClient } from '../src/client.js';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p: string) => JSON.parse(readFileSync(join(root, p), 'utf8'));
const pkg = read('package.json');

describe('version sync', () => {
  it('keeps every version marker equal to package.json', () => {
    expect(versionSyncTest({ srcDir: join(root, 'src'), pkgPath: join(root, 'package.json') })).toEqual(
      [],
    );
  });

  it('keeps the manifests on the same version', () => {
    const v = pkg.version;
    expect(read('manifest.json').version).toBe(v);
    expect(read('server.json').version).toBe(v);
    expect(read('server.json').packages[0].version).toBe(v);
    expect(read('.claude-plugin/plugin.json').version).toBe(v);
    expect(read('.claude-plugin/marketplace.json').metadata.version).toBe(v);
    expect(read('.claude-plugin/marketplace.json').plugins[0].version).toBe(v);
  });

  it('registers every version-bearing file with release-please', () => {
    const extras = read('release-please-config.json').packages['.']['extra-files'];
    const paths = extras.map((e: unknown) => (typeof e === 'string' ? e : (e as { path: string }).path));
    for (const p of [
      'manifest.json',
      'server.json',
      '.claude-plugin/plugin.json',
      '.claude-plugin/marketplace.json',
      'src/version.ts',
    ]) {
      expect(paths).toContain(p);
    }
  });
});

describe('publish scaffold', () => {
  it('declares the repository url npm provenance validates against', () => {
    // Without this the publish 422s AFTER release-please has already tagged,
    // and re-running the job cannot fix it.
    expect(pkg.repository?.url).toBe(
      'git+https://github.com/chrischall/office-outlook-mcp.git',
    );
  });

  it('publishes under the chrischall scope with public access', () => {
    // Scoped because `outlook-mcp` and `office-mcp` both exist on npm, and an
    // unscoped `office-outlook-mcp` risks npm's "too similar" rejection at
    // publish time. Scoped names are exempt from that check.
    expect(pkg.name).toBe('@chrischall/office-outlook-mcp');
    expect(pkg.publishConfig?.access).toBe('public');
  });

  it('ships the files an install and a registration need', () => {
    // `skills` and `mint.yaml` are both silent-omission traps: without them the
    // access skill never ships and an --npm registration reads a blank wizard.
    for (const f of ['dist', 'skills', 'mint.yaml', 'server.json', '.claude-plugin']) {
      expect(pkg.files).toContain(f);
    }
  });

  it('keeps the registry description within the 100-char schema limit', () => {
    expect(read('server.json').description.length).toBeLessThanOrEqual(100);
  });

  it('keeps the mcpb runtime floor on an LTS Node so LTS users can install', () => {
    expect(read('manifest.json').runtimes.node).toBe('>=22.5');
  });
});

describe('manifest tool roster', () => {
  it('matches the registered tools in BOTH directions', async () => {
    // A tool missing from manifest.json is invisible to an mcpb host even
    // though the server answers it, and nothing else reads that file.
    const client = {
      get: async () => ({ value: [] }),
      write: async () => ({}),
      refreshTokenIfNeeded: async () => {},
      tokenExpiresAt: () => null,
      tokenSource: 'test',
      apiBase: 'https://outlook.office.com/api/v2.0',
    } as unknown as OutlookClient;

    const h = await createTestHarness((server: McpServer) => {
      registerMailTools(server, client);
      registerCalendarTools(server, client);
      registerDirectoryTools(server, client);
      registerWriteTools(server, client);
      registerHealthcheckTool(server, client);
    });
    const registered = (await h.listTools()).map((t) => t.name).sort();
    await h.close();

    const declared = read('manifest.json').tools.map((t: { name: string }) => t.name).sort();
    expect(declared).toEqual(registered);

    for (const t of read('manifest.json').tools) {
      expect(t.description, `${t.name} needs a description`).toBeTruthy();
    }
  });
});
