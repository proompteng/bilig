import { afterEach, describe, expect, it, vi } from 'vitest'

afterEach(() => {
  vi.resetModules()
  vi.doUnmock('@bilig/wasm-kernel')
  vi.clearAllMocks()
})

describe('WasmKernelFacade init failures', () => {
  it('can initialize synchronously in Node-like runtimes', async () => {
    const { WasmKernelFacade } = await import('../wasm-facade.js')
    const facade = new WasmKernelFacade()

    expect(facade.initSyncIfPossible()).toBe(true)
    expect(facade.ready).toBe(true)
  })

  it('keeps the facade unready when kernel creation fails', async () => {
    vi.doMock('@bilig/wasm-kernel', () => ({
      createKernel: vi.fn(async () => {
        throw new Error('kernel init failed')
      }),
      createKernelSync: vi.fn(() => {
        throw new Error('kernel init failed')
      }),
    }))

    const { WasmKernelFacade } = await import('../wasm-facade.js')
    const facade = new WasmKernelFacade()

    await expect(facade.init()).resolves.toBeUndefined()
    expect(facade.ready).toBe(false)
  })
})
