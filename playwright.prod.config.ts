import { defineConfig, devices } from "@playwright/test";

const baseURL = process.env["BILIG_PROD_BASE_URL"] ?? "https://bilig.proompteng.ai";

export default defineConfig({
  testDir: "./e2e/tests",
  testMatch: "**/prod-smoke.pw.ts",
  fullyParallel: false,
  retries: 0,
  timeout: 60_000,
  expect: {
    timeout: 15_000,
  },
  use: {
    ...devices["Desktop Chrome"],
    baseURL,
    trace: "retain-on-failure",
    video: "retain-on-failure",
    screenshot: "only-on-failure",
  },
});
