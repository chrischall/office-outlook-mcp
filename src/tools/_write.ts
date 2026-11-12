/**
 * Glue between this server's write tools and mcp-utils' batched writes
 * (`registerWrite` / `registerBatchTool`, docs/CONFIRM-TOKEN.md §14).
 *
 * `registerWrite` builds the `registerTool` config itself and takes no
 * `_meta`, and it always gates. Both helpers here patch `registerTool` on the
 * SAME server object for the duration of one call (not a proxy), because
 * registerWrite's per-server write registry — what `outlook_batch` reads — is
 * keyed on the real server. The original is restored afterwards, even when the
 * registration throws.
 */
import type { CallToolResult, McpServer, ServerContext } from '@modelcontextprotocol/server';
import { registerWrite, type WriteDefinition } from '@chrischall/mcp-utils';
import type { z } from 'zod';

type Register = (...args: unknown[]) => unknown;

function patchRegisterTool(server: McpServer, replace: (original: Register) => Register, run: () => void): void {
  const target = server as unknown as { registerTool: Register };
  const hadOwn = Object.prototype.hasOwnProperty.call(target, 'registerTool');
  const original = target.registerTool;
  target.registerTool = replace(original);
  try {
    run();
  } finally {
    if (hadOwn) target.registerTool = original;
    else delete (target as { registerTool?: unknown }).registerTool;
  }
}

/**
 * Run `register` with `_meta` (the mint hint) added to every tool it
 * registers — for mcp-utils' `registerWrite`, `registerConfirmTool` and
 * `registerBatchTool`, which take no `_meta`.
 */
export function withToolMeta(server: McpServer, meta: Record<string, unknown>, register: () => void): void {
  patchRegisterTool(
    server,
    (original) => (name, config, ...rest) =>
      original.call(server, name, { ...(config as Record<string, unknown>), _meta: meta }, ...rest),
    register,
  );
}

/** `registerWrite`, with `_meta` (the mint hint) added to the tool it registers. */
export function registerGatedWrite<S extends z.ZodObject, State, Result>(
  server: McpServer,
  meta: Record<string, unknown>,
  def: WriteDefinition<S, State, Result>,
): void {
  withToolMeta(server, meta, () => registerWrite(server, def));
}

/**
 * A write that is NOT gated as a single tool (nothing reaches anyone else) but
 * may still be an op in `outlook_batch`, where the batch's one approval covers
 * it. Its definition goes into the batch registry only — registerWrite's own
 * (gated) tool is swallowed — and the single tool runs plan → execute → render
 * with no confirmation, as it did before.
 */
export function registerUngatedWrite<S extends z.ZodObject, State, Result>(
  server: McpServer,
  config: { description: string; meta: Record<string, unknown> },
  def: WriteDefinition<S, State, Result> & { render(result: Result): CallToolResult },
): void {
  patchRegisterTool(server, () => () => undefined, () => registerWrite(server, def));
  // Through the untyped signature: the SDK's overloads cannot see that a
  // generic `S extends z.ZodObject` is a valid input schema.
  (server as unknown as { registerTool: Register }).registerTool(
    def.name,
    {
      description: config.description,
      annotations: def.annotations,
      _meta: config.meta,
      inputSchema: def.inputSchema,
    },
    async (args: z.infer<S>, ctx: ServerContext) => {
      const { state } = await def.plan(args, ctx);
      return def.render(await def.execute(state, ctx));
    },
  );
}
