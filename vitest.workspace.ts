import { fileURLToPath } from "node:url";
import { defineWorkspace } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));
const alias = {
  "@bilig/protocol": `${rootDir}packages/protocol/src/index.ts`,
  "@bilig/formula": `${rootDir}packages/formula/src/index.ts`,
  "@bilig/formula/program-arena": `${rootDir}packages/formula/src/program-arena.ts`,
  "@bilig/core": `${rootDir}packages/core/src/index.ts`,
  "@bilig/crdt": `${rootDir}packages/crdt/src/index.ts`,
  "@bilig/binary-protocol": `${rootDir}packages/binary-protocol/src/index.ts`,
  "@bilig/agent-api": `${rootDir}packages/agent-api/src/index.ts`,
  "@bilig/storage-browser": `${rootDir}packages/storage-browser/src/index.ts`,
  "@bilig/storage-server": `${rootDir}packages/storage-server/src/index.ts`,
  "@bilig/wasm-kernel": `${rootDir}packages/wasm-kernel/src/index.ts`,
  "@bilig/worker-transport": `${rootDir}packages/worker-transport/src/index.ts`,
  "@bilig/renderer": `${rootDir}packages/renderer/src/index.ts`,
  "@bilig/grid": `${rootDir}packages/grid/src/index.ts`
};

export default defineWorkspace([
  {
    resolve: {
      alias
    },
    test: {
      name: "packages",
      include: ["packages/*/src/**/*.test.ts", "packages/*/src/**/*.test.tsx"],
      exclude: ["**/dist/**", "**/build/**"],
      environment: "node",
      coverage: {
        provider: "v8",
        reporter: ["text", "lcov", "json-summary"],
        include: [
          "packages/core/src/**/*.ts",
          "packages/formula/src/**/*.ts",
          "packages/renderer/src/**/*.ts"
        ],
        exclude: [
          "**/__tests__/**",
          "**/*.d.ts",
          "packages/core/src/index.ts",
          "packages/core/src/snapshot.ts",
          "packages/formula/src/index.ts",
          "packages/formula/src/ast.ts",
          "packages/renderer/src/index.ts"
        ],
        thresholds: {
          lines: 85,
          statements: 85,
          functions: 85,
          branches: 70
        }
      }
    }
  },
  {
    resolve: {
      alias
    },
    test: {
      name: "apps-node",
      include: [
        "apps/local-server/src/**/*.test.ts",
        "apps/local-server/src/**/*.test.tsx",
        "apps/sync-server/src/**/*.test.ts",
        "apps/sync-server/src/**/*.test.tsx"
      ],
      exclude: ["**/dist/**", "**/build/**"],
      environment: "node"
    }
  },
  {
    resolve: {
      alias
    },
    test: {
      name: "apps-jsdom",
      include: [
        "apps/playground/src/**/*.test.ts",
        "apps/playground/src/**/*.test.tsx",
        "apps/web/src/**/*.test.ts",
        "apps/web/src/**/*.test.tsx"
      ],
      exclude: ["**/dist/**", "**/build/**"],
      environment: "jsdom"
    }
  }
]);
