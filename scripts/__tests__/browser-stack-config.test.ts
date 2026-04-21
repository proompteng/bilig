import { describe, expect, it } from 'vitest'
import { buildBrowserLocalStackCommand, resolveBrowserLocalWebMode } from '../browser-stack-config.js'

describe('browser stack config', () => {
  it('defaults browser test web mode to preview', () => {
    expect(resolveBrowserLocalWebMode({})).toBe('preview')
  })

  it('allows opting back into dev mode explicitly', () => {
    expect(resolveBrowserLocalWebMode({ BILIG_BROWSER_WEB_MODE: 'dev' })).toBe('dev')
  })

  it('adds the resolved web mode to the local playwright stack command', () => {
    const command = buildBrowserLocalStackCommand({
      browserWebPort: '4180',
      browserAppPort: '54422',
      browserPostgresPort: '55433',
      browserZeroPort: '54849',
      disableCompose: true,
      remoteSyncEnabled: false,
      webMode: 'preview',
    })

    expect(command).toContain('BILIG_DEV_WEB_SERVER_MODE=preview')
    expect(command).toContain('BILIG_DEV_APP_SERVER_MODE=run')
  })
})
