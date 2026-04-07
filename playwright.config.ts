import { defineConfig, devices } from "@playwright/test";

const browserStack = process.env["BILIG_BROWSER_STACK"];
const useComposeBrowserStack = browserStack === "compose";
const fuzzBrowserMode = process.env["BILIG_FUZZ_BROWSER"] === "1";
const ciContainerMode =
  process.platform === "linux" && (process.env["CI"] === "1" || process.env["CI"] === "true");
const browserWebPort = process.env["BILIG_E2E_WEB_PORT"] ?? "4180";
const browserAppPort = process.env["BILIG_E2E_SYNC_SERVER_PORT"] ?? "54422";
const browserZeroPort = process.env["BILIG_E2E_ZERO_PORT"] ?? "54849";
const browserPostgresPort = process.env["BILIG_E2E_POSTGRES_PORT"] ?? "55433";
const browserBaseUrl = process.env["BILIG_E2E_BASE_URL"] ?? `http://127.0.0.1:${browserWebPort}`;
const browserReadyUrl =
  process.env["BILIG_E2E_READY_URL"] ?? `http://127.0.0.1:${browserAppPort}/healthz`;
const browserLocalStackCommand =
  `BILIG_WEB_DEV_PORT=${browserWebPort} ` +
  `PORT=${browserAppPort} ` +
  `BILIG_DEV_POSTGRES_PORT=${browserPostgresPort} ` +
  `BILIG_DEV_ZERO_PORT=${browserZeroPort} ` +
  `BILIG_ZERO_PROXY_UPSTREAM=http://127.0.0.1:${browserZeroPort} ` +
  "BILIG_DEV_COMPOSE_PROJECT=bilig-playwright-local " +
  "BILIG_DEV_CLEANUP_COMPOSE=true " +
  "bun scripts/run-dev-web-local.ts";
const chromiumLaunchArgs = ciContainerMode
  ? ["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
  : ["--enable-unsafe-webgpu", "--ignore-gpu-blocklist"];

export default defineConfig({
  testDir: "./e2e/tests",
  testMatch: "**/web-shell.pw.ts",
  fullyParallel: false,
  retries: 0,
  timeout: fuzzBrowserMode ? 600_000 : 30_000,
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
          args: chromiumLaunchArgs,
        },
      },
    },
  ],
  ...(useComposeBrowserStack
    ? {}
    : {
        webServer: {
          command: browserLocalStackCommand,
          url: browserReadyUrl,
          reuseExistingServer: false,
          timeout: 300_000,
        },
      }),
});
