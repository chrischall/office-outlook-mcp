import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Pins the token cache into a temp dir and turns it off, so no test can
    // reach the developer's real ~/.office-outlook-mcp — see tests/_setup.ts.
    setupFiles: ['./tests/_setup.ts'],
    coverage: {
      provider: 'v8',
      // The fleet's measured set: source only, minus the stdio entrypoint.
      include: ['src/**/*.ts'],
      exclude: ['src/index.ts'],
      reporter: ['text', 'html'],
      /**
       * A RATCHET, not the fleet's 100%.
       *
       * Measured at 89.6 / 86.3 / 89.0 / 89.9 across 92 tests; these sit just
       * under that so ordinary measurement noise does not fail a build.
       * RAISE THEM as coverage improves.
       *
       * What remains uncovered is uncovered for a reason worth stating:
       * `captureTokenLazily` in client.ts and the capture race in
       * auth-fetchproxy.ts both need a live browser bridge, which this suite
       * deliberately cannot have — tests/_setup.ts exists to guarantee it
       * never does. Those paths were verified by hand against a real mailbox;
       * see docs/OUTLOOK-API.md.
       */
      thresholds: { statements: 89, branches: 86, functions: 88, lines: 89 },
    },
  },
});
