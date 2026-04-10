import type { KnipConfig } from "knip";

const config: KnipConfig = {
  entry: [
    "playwright.config.ts",
    "playwright.prod.config.ts",
    "vitest.config.ts",
    "vitest.workspace.ts",
    "scripts/**/*.ts",
    "e2e/tests/**/*.pw.ts",
    "apps/*/src/index.ts",
    "apps/*/src/main.tsx",
    "apps/*/src/**/*.test.ts",
    "apps/*/src/**/*.test.tsx",
    "packages/*/src/index.ts",
    "packages/*/src/**/*.test.ts",
    "packages/*/src/**/*.test.tsx",
    "packages/wasm-kernel/scripts/**/*.ts",
  ],
  project: [
    "apps/**/*.{ts,tsx}",
    "packages/**/*.{ts,tsx}",
    "scripts/**/*.ts",
    "e2e/**/*.{ts,tsx}",
    "!**/dist/**",
    "!**/build/**",
  ],
  ignore: [
    "packages/wasm-kernel/assembly/**",
    "apps/bilig/src/recalc/worker.ts",
    "e2e/tests/prod-smoke.pw.ts",
  ],
  ignoreDependencies: [
    "@effect/platform",
    "@effect/platform-node",
    "@types/react-dom",
    "@types/ws",
    "@xstate/react",
    "assemblyscript",
    "effect",
    "xstate",
  ],
  workspaces: {
    "apps/web": {},
    "packages/core": {},
    "packages/formula": {},
    "packages/grid": {},
    "packages/runtime-kernel": {},
    "packages/worker-transport": {},
  },
};

export default config;
