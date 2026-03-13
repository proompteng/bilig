import { defineWorkspace } from "vitest/config";

export default defineWorkspace([
  {
    test: {
      name: "packages",
      include: ["packages/**/*.test.ts"],
      environment: "node",
      coverage: {
        provider: "v8",
        reporter: ["text", "lcov"]
      }
    }
  },
  {
    test: {
      name: "playground",
      include: ["apps/playground/**/*.test.ts", "apps/playground/**/*.test.tsx"],
      environment: "jsdom"
    }
  }
]);
