export type BrowserStack = 'auto' | 'compose' | 'local'
export type BrowserLocalWebMode = 'dev' | 'preview'

export function resolveBrowserStack(env: { BILIG_BROWSER_STACK?: string | undefined }): BrowserStack {
  const value = env.BILIG_BROWSER_STACK
  if (value === undefined || value === 'auto') {
    return 'auto'
  }
  if (value === 'compose' || value === 'local') {
    return value
  }

  throw new Error(`BILIG_BROWSER_STACK must be "auto", "compose", or "local", got ${value}`)
}

export function resolveBrowserLocalWebMode(env: { BILIG_BROWSER_WEB_MODE?: string | undefined }): BrowserLocalWebMode {
  const value = env.BILIG_BROWSER_WEB_MODE
  if (value === undefined || value === 'preview') {
    return 'preview'
  }
  if (value === 'dev') {
    return 'dev'
  }

  throw new Error(`BILIG_BROWSER_WEB_MODE must be "preview" or "dev", got ${value}`)
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
