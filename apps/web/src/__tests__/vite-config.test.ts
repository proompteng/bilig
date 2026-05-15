import { describe, expect, it } from 'vitest'
import webViteConfig, { crossOriginIsolationHeaders } from '../../vite.config'

describe('web vite config', () => {
  it('serves dev and preview with cross-origin isolation headers', () => {
    expect(webViteConfig.server?.headers).toEqual(crossOriginIsolationHeaders)
    expect(webViteConfig.preview?.headers).toEqual(crossOriginIsolationHeaders)
  })

  it('keeps worker-only dynamic imports in split module worker chunks', () => {
    expect(webViteConfig.worker?.format).toBe('es')
    expect(webViteConfig.worker?.rolldownOptions?.output).toEqual(webViteConfig.build?.rolldownOptions?.output)
  })
})
