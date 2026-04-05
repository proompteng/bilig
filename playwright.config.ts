import { defineConfig, devices } from "@playwright/test";

const browserStack = process.env["BILIG_BROWSER_STACK"];
const useComposeBrowserStack = browserStack === "compose";
const browserBaseUrl = process.env["BILIG_E2E_BASE_URL"] ?? "http://127.0.0.1:4180";
const browserReadyUrl = process.env["BILIG_E2E_READY_URL"] ?? "http://127.0.0.1:4181/healthz";

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
        launchOptions: {
          args: ["--enable-unsafe-webgpu", "--ignore-gpu-blocklist"],
        },
      },
    },
  ],
  ...(useComposeBrowserStack
    ? {}
    : {
        webServer: {
          command: "bun scripts/run-playwright-stack.ts",
          url: browserReadyUrl,
          reuseExistingServer: false,
          timeout: 300_000,
        },
      }),
});
