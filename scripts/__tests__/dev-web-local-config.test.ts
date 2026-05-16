import { describe, expect, it } from 'vitest'

import { resolveDevAppServerMode, resolveDevWebServerMode } from '../dev-web-local-config.js'

describe('dev web local config', () => {
  it('defaults local web server mode to dev and accepts explicit modes', () => {
    expect(resolveDevWebServerMode({})).toBe('dev')
    expect(resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'dev' })).toBe('dev')
    expect(resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'preview' })).toBe('preview')
  })

  it('rejects malformed local web server modes', () => {
    expect(() => resolveDevWebServerMode({ BILIG_DEV_WEB_SERVER_MODE: 'prevew' })).toThrow(
      'BILIG_DEV_WEB_SERVER_MODE must be "dev" or "preview", got prevew',
    )
  })

  it('defaults local app server mode to watch and accepts explicit modes', () => {
    expect(resolveDevAppServerMode({})).toBe('watch')
    expect(resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'watch' })).toBe('watch')
    expect(resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'run' })).toBe('run')
  })

  it('rejects malformed local app server modes', () => {
    expect(() => resolveDevAppServerMode({ BILIG_DEV_APP_SERVER_MODE: 'rn' })).toThrow(
      'BILIG_DEV_APP_SERVER_MODE must be "watch" or "run", got rn',
    )
  })
})
