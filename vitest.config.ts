import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  resolve: {
    alias: [
      {
        find: "@bilig/formula/program-arena",
        replacement: `${rootDir}packages/formula/src/program-arena.ts`,
      },
      { find: "@bilig/protocol", replacement: `${rootDir}packages/protocol/src/index.ts` },
      { find: "@bilig/formula", replacement: `${rootDir}packages/formula/src/index.ts` },
      { find: "@bilig/core", replacement: `${rootDir}packages/core/src/index.ts` },
      { find: "@bilig/crdt", replacement: `${rootDir}packages/crdt/src/index.ts` },
      {
        find: "@bilig/binary-protocol",
        replacement: `${rootDir}packages/binary-protocol/src/index.ts`,
      },
      { find: "@bilig/agent-api", replacement: `${rootDir}packages/agent-api/src/index.ts` },
      {
        find: "@bilig/storage-browser",
        replacement: `${rootDir}packages/storage-browser/src/index.ts`,
      },
      { find: "@bilig/zero-sync", replacement: `${rootDir}packages/zero-sync/src/index.ts` },
      {
        find: "@bilig/storage-server",
        replacement: `${rootDir}packages/storage-server/src/index.ts`,
      },
      { find: "@bilig/wasm-kernel", replacement: `${rootDir}packages/wasm-kernel/src/index.ts` },
      {
        find: "@bilig/worker-transport",
        replacement: `${rootDir}packages/worker-transport/src/index.ts`,
      },
      { find: "@bilig/renderer", replacement: `${rootDir}packages/renderer/src/index.ts` },
      { find: "@bilig/grid", replacement: `${rootDir}packages/grid/src/index.ts` },
      { find: "@bilig/test-fuzz", replacement: `${rootDir}packages/test-fuzz/src/index.ts` },
    ],
  },
  test: {
    environment: "node",
    include: [
      "packages/*/src/**/*.test.ts",
      "packages/*/src/**/*.test.tsx",
      "apps/*/src/**/*.test.ts",
      "apps/*/src/**/*.test.tsx",
    ],
    exclude: ["**/dist/**", "**/build/**"],
    coverage: {
      provider: "v8",
      reporter: ["text", "lcov", "json", "json-summary"],
      include: [
        "packages/core/src/**/*.ts",
        "packages/formula/src/**/*.ts",
        "packages/renderer/src/**/*.ts",
      ],
      exclude: [
        "**/__tests__/**",
        "**/*.d.ts",
        "packages/core/src/index.ts",
        "packages/core/src/snapshot.ts",
        "packages/formula/src/index.ts",
        "packages/formula/src/ast.ts",
        "packages/renderer/src/index.ts",
      ],
      thresholds: {
        lines: 90,
        statements: 90,
        functions: 90,
        branches: 70,
      },
    },
  },
});
