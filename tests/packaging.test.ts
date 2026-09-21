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

describe('opencode install docs', () => {
  const readme = readFileSync(join(root, 'README.md'), 'utf8');
  const mcpBlocks = [...readme.matchAll(/```jsonc?\n([\s\S]*?)```/g)]
    .map((m) => m[1])
    .filter((b) => b.includes('"mcp"'));

  it('documents opencode at all', () => {
    expect(readme).toMatch(/opencode/i);
    expect(mcpBlocks.length).toBeGreaterThan(0);
  });

  it('leads with the nested `servers` shape, which serves opencode 2 AND current 1', () => {
    // Measured 2026-09-20 against real binaries: opencode 2.0.11 reads
    // `mcp.servers.<name>`; opencode 1.18.28 and 1.18.31 read it too. Only
    // older 1.x rejects it, and does so LOUDLY ("Configuration is invalid"),
    // which is the failure mode you want — the user is told to switch shapes.
    const primary = mcpBlocks[0];
    expect(primary).toContain('"servers"');
    expect(primary).toContain('@chrischall/office-outlook-mcp');
  });

  it('never shows both shapes in one file — opencode 2 drops the lot in silence', () => {
    // THE trap, and the reason this test exists. Given a `mcp` block carrying
    // both a v1-style `mcp.<name>` entry and `mcp.servers`, opencode 2.0.11
    // parses the document (it shows up in `debug config`) and then reports "No
    // MCP servers configured" — no error, no warning, every server gone.
    // Verified A/B in one directory against one service: remove the v1 sibling
    // and the same file connects.
    for (const block of mcpBlocks) {
      if (!block.includes('"servers"')) continue;
      const mcp = (JSON.parse(block) as { mcp: Record<string, unknown> }).mcp;
      expect(Object.keys(mcp).filter((k) => k !== 'servers' && k !== 'timeout')).toEqual([]);
    }
  });

  it('keeps every documented block valid JSON', () => {
    for (const block of mcpBlocks) expect(() => JSON.parse(block)).not.toThrow();
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

  it('points the plugin at a config that resolves under a plugin install', () => {
    // The two launch paths need DIFFERENT configs, which is why they are two
    // files: a plugin install defines `${CLAUDE_PLUGIN_ROOT}` and a
    // project-scoped `.mcp.json` does not. Sharing one file broke whichever
    // path it was not written for.
    const pluginMcp = read('.claude-plugin/plugin.json').mcp as string;
    const cfg = read(join('.claude-plugin', pluginMcp.replace(/^\.\//, '')));
    const server = cfg.mcpServers.outlook;
    expect(server.args.join(' ')).toContain('${CLAUDE_PLUGIN_ROOT}');
    expect(server.args.join(' ')).toContain('dist/bundle.js');
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
    // Nested under `compatibility` — there is NO top-level `runtimes` in the
    // mcpb schema, and `mcpb pack` rejects the whole manifest for an
    // unrecognised key. That failure lands in the PUBLISH job, after
    // release-please has already tagged and cut the GitHub Release, and it
    // runs BEFORE `npm publish` — so one stray key means the tag exists, the
    // Release exists, and npm never moves.
    expect(read('manifest.json').compatibility.runtimes.node).toBe('>=22.5.0');
  });

  it('carries no key the mcpb schema would reject', () => {
    // A value-only assertion cannot see a key in the wrong PLACE, which is how
    // the invalid manifest shipped green. This pins the shape instead.
    const allowed = new Set([
      '$schema', 'manifest_version', 'name', 'display_name', 'version',
      'description', 'author', 'repository', 'homepage', 'support', 'license',
      'keywords', 'server', 'user_config', 'tools', 'compatibility', 'icon',
      'screenshots', 'long_description', 'documentation', 'privacy_policies',
      'tools_generated', 'prompts', 'prompts_generated',
    ]);
    const unknown = Object.keys(read('manifest.json')).filter((k) => !allowed.has(k));
    expect(unknown).toEqual([]);
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
