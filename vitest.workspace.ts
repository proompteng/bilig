import { fileURLToPath } from "node:url";
import { defineWorkspace } from "vitest/config";

const rootDir = fileURLToPath(new URL(".", import.meta.url));
const alias = {
  "@bilig/protocol": `${rootDir}packages/protocol/src/index.ts`,
  "@bilig/formula": `${rootDir}packages/formula/src/index.ts`,
  "@bilig/core": `${rootDir}packages/core/src/index.ts`,
  "@bilig/crdt": `${rootDir}packages/crdt/src/index.ts`,
  "@bilig/wasm-kernel": `${rootDir}packages/wasm-kernel/src/index.ts`
};

export default defineWorkspace([
  {
    resolve: {
      alias
    },
    test: {
      name: "packages",
      include: ["packages/*/src/**/*.test.ts"],
      exclude: ["**/dist/**", "**/build/**"],
      environment: "node",
      coverage: {
        provider: "v8",
        reporter: ["text", "lcov"]
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
