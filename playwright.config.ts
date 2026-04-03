import { defineConfig, devices } from "@playwright/test";

const browserStack = process.env["BILIG_BROWSER_STACK"];
const useComposeBrowserStack = browserStack === "compose";
const browserBaseUrl = process.env["BILIG_E2E_BASE_URL"] ?? "http://127.0.0.1:4180";

export default defineConfig({
  testDir: "./e2e/tests",
  testMatch: "**/web-shell.pw.ts",
  fullyParallel: false,
  retries: 0,
  reporter: "list",
  use: {
    baseURL: browserBaseUrl,
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
  ...(useComposeBrowserStack
    ? {}
    : {
        webServer: {
          command: "bun scripts/run-playwright-stack.ts",
          port: 4180,
          reuseExistingServer: false,
          timeout: 300_000,
        },
      }),
});
