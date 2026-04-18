import { describe, expect, it } from 'vitest'
import { ensureWasmBinaryPathForNode } from '../index.js'

const fileURLToPath = (url: URL): string => url.pathname

describe('wasm kernel binary resolution', () => {
  const importMetaUrl = 'file:///repo/packages/wasm-kernel/src/index.ts'

  it('returns the existing release.wasm path', () => {
    const wasmPath = ensureWasmBinaryPathForNode({
      importMetaUrl,
      existsSync: (path) => path === '/repo/packages/wasm-kernel/build/release.wasm',
      fileURLToPath,
    })

    expect(wasmPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
  })

  it('throws a clear error when the artifact is missing', () => {
    expect(() =>
      ensureWasmBinaryPathForNode({
        importMetaUrl,
        existsSync: () => false,
        fileURLToPath,
      }),
    ).toThrow("Run 'pnpm wasm:build'")
  })
})
