export type BrowserLocalWebMode = 'dev' | 'preview'

export function resolveBrowserLocalWebMode(env: { BILIG_BROWSER_WEB_MODE?: string | undefined }): BrowserLocalWebMode {
  return env.BILIG_BROWSER_WEB_MODE === 'dev' ? 'dev' : 'preview'
}

export function buildBrowserLocalStackCommand(input: {
  browserWebPort: string
  browserAppPort: string
  browserPostgresPort: string
  browserZeroPort: string
  disableCompose: boolean
  remoteSyncEnabled: boolean
  webMode: BrowserLocalWebMode
}): string {
  return [
    `BILIG_WEB_DEV_PORT=${input.browserWebPort}`,
    `PORT=${input.browserAppPort}`,
    `BILIG_DEV_POSTGRES_PORT=${input.browserPostgresPort}`,
    `BILIG_DEV_ZERO_PORT=${input.browserZeroPort}`,
    `BILIG_DEV_WEB_SERVER_MODE=${input.webMode}`,
    'BILIG_DEV_APP_SERVER_MODE=run',
    'BILIG_DEV_COMPOSE_PROJECT=bilig-playwright-local',
    'BILIG_DEV_CLEANUP_COMPOSE=true',
    input.disableCompose ? 'BILIG_DEV_DISABLE_COMPOSE=1' : null,
    input.remoteSyncEnabled ? null : 'BILIG_E2E_REMOTE_SYNC=0',
    'bun scripts/run-dev-web-local.ts',
  ]
    .filter((segment): segment is string => segment !== null)
    .join(' ')
}
