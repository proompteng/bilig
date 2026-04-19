import { defineConfig, devices } from '@playwright/test'
import { buildBrowserLocalStackCommand, resolveBrowserLocalWebMode } from './scripts/browser-stack-config.js'

const browserStack = process.env['BILIG_BROWSER_STACK']
const useComposeBrowserStack = browserStack === 'compose'
const fuzzBrowserMode = process.env['BILIG_FUZZ_BROWSER'] === '1'
const localNoCompose = process.env['BILIG_DEV_DISABLE_COMPOSE'] === '1'
const remoteSyncEnabled = process.env['BILIG_E2E_REMOTE_SYNC'] !== '0'
const ciContainerMode = process.platform === 'linux' && (process.env['CI'] === '1' || process.env['CI'] === 'true')
const browserHost = process.env['BILIG_E2E_HOST'] ?? '127.0.0.1'
const browserWebPort = process.env['BILIG_E2E_WEB_PORT'] ?? '4180'
const browserAppPort = process.env['BILIG_E2E_SYNC_SERVER_PORT'] ?? '54422'
const browserZeroPort = process.env['BILIG_E2E_ZERO_PORT'] ?? '54849'
const browserPostgresPort = process.env['BILIG_E2E_POSTGRES_PORT'] ?? '55433'
const browserBaseUrl = process.env['BILIG_E2E_BASE_URL'] ?? `http://${browserHost}:${browserWebPort}`
const browserReadyUrl =
  process.env['BILIG_E2E_READY_URL'] ?? (remoteSyncEnabled ? `${browserBaseUrl}/zero/keepalive` : `${browserBaseUrl}/runtime-config.json`)
const browserLocalStackCommand = buildBrowserLocalStackCommand({
  browserWebPort,
  browserAppPort,
  browserPostgresPort,
  browserZeroPort,
  disableCompose: localNoCompose,
  remoteSyncEnabled,
  webMode: resolveBrowserLocalWebMode(process.env),
})
const chromiumLaunchArgs = ciContainerMode
  ? ['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu']
  : ['--enable-unsafe-webgpu', '--ignore-gpu-blocklist']

export default defineConfig({
  testDir: './e2e/tests',
  testMatch: '**/web-shell*.pw.ts',
  fullyParallel: false,
  retries: 0,
  timeout: fuzzBrowserMode ? 600_000 : 30_000,
  reporter: 'list',
  use: {
    baseURL: browserBaseUrl,
    trace: 'retain-on-failure',
  },
  projects: [
    {
      name: 'chromium',
      use: {
        ...devices['Desktop Chrome'],
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
})
