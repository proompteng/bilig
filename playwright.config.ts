import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "./e2e/tests",
  testMatch: "**/playground.pw.ts",
  fullyParallel: false,
  retries: 0,
  reporter: "list",
  use: {
    baseURL: "http://127.0.0.1:4179",
    trace: "retain-on-failure",
  },
  projects: [
    {
      name: "chromium",
      use: {
        ...devices["Desktop Chrome"],
      },
    },
  ],
  webServer: {
    command: "pnpm --filter @bilig/playground preview:ci",
    port: 4179,
    reuseExistingServer: false,
    timeout: 120_000,
  },
});
