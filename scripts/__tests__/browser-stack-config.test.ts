import { describe, expect, it } from 'vitest'
import { buildBrowserLocalStackCommand, resolveBrowserLocalWebMode, resolveBrowserStack } from '../browser-stack-config.js'

describe('browser stack config', () => {
  it('defaults browser test stack selection to auto and accepts explicit stacks', () => {
    expect(resolveBrowserStack({})).toBe('auto')
    expect(resolveBrowserStack({ BILIG_BROWSER_STACK: 'auto' })).toBe('auto')
    expect(resolveBrowserStack({ BILIG_BROWSER_STACK: 'compose' })).toBe('compose')
    expect(resolveBrowserStack({ BILIG_BROWSER_STACK: 'local' })).toBe('local')
  })

  it('rejects malformed browser stack selections instead of silently changing coverage', () => {
    expect(() => resolveBrowserStack({ BILIG_BROWSER_STACK: 'docker' })).toThrow(
      'BILIG_BROWSER_STACK must be "auto", "compose", or "local", got docker',
    )
  })

  it('defaults browser test web mode to preview', () => {
    expect(resolveBrowserLocalWebMode({})).toBe('preview')
  })

  it('allows opting back into dev mode explicitly', () => {
    expect(resolveBrowserLocalWebMode({ BILIG_BROWSER_WEB_MODE: 'preview' })).toBe('preview')
    expect(resolveBrowserLocalWebMode({ BILIG_BROWSER_WEB_MODE: 'dev' })).toBe('dev')
  })

  it('rejects malformed browser web mode selections', () => {
    expect(() => resolveBrowserLocalWebMode({ BILIG_BROWSER_WEB_MODE: 'prevew' })).toThrow(
      'BILIG_BROWSER_WEB_MODE must be "preview" or "dev", got prevew',
    )
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
