import { fileURLToPath } from "node:url";
import { defineWorkspace } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));
const alias = {
  "@bilig/protocol": `${rootDir}packages/protocol/src/index.ts`,
  "@bilig/formula": `${rootDir}packages/formula/src/index.ts`,
  "@bilig/core": `${rootDir}packages/core/src/index.ts`,
  "@bilig/crdt": `${rootDir}packages/crdt/src/index.ts`,
  "@bilig/wasm-kernel": `${rootDir}packages/wasm-kernel/src/index.ts`,
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
        reporter: ["text", "lcov"],
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
          lines: 80,
          statements: 80,
          functions: 85,
          branches: 65
        }
      }
    }
  },
  {
    resolve: {
      alias
    },
    test: {
      name: "playground",
      include: ["apps/playground/src/**/*.test.ts", "apps/playground/src/**/*.test.tsx"],
      exclude: ["**/dist/**", "**/build/**"],
      environment: "jsdom"
    }
  }
]);
