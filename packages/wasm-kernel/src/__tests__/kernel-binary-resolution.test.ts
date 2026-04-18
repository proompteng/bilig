import { describe, expect, it, vi } from 'vitest'
import { ensureWasmBinaryPathForNode } from '../index.js'

const fileURLToPath = (url: URL): string => url.pathname

describe('wasm kernel binary resolution', () => {
  const importMetaUrl = 'file:///repo/packages/wasm-kernel/src/index.ts'

  it('returns the existing release.wasm path without building', () => {
    const runBuildSync = vi.fn()
    const wasmPath = ensureWasmBinaryPathForNode({
      importMetaUrl,
      existsSync: (path) => path === '/repo/packages/wasm-kernel/build/release.wasm',
      fileURLToPath,
      runBuildSync,
    })

    expect(wasmPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(runBuildSync).not.toHaveBeenCalled()
  })

  it('builds the wasm artifact on demand when the build script exists', () => {
    const existingPaths = new Set<string>(['/repo/packages/wasm-kernel/scripts/build.ts'])
    const runBuildSync = vi.fn((packageRootPath: string) => {
      expect(packageRootPath).toBe('/repo/packages/wasm-kernel/')
      existingPaths.add('/repo/packages/wasm-kernel/build/release.wasm')
    })

    const wasmPath = ensureWasmBinaryPathForNode({
      importMetaUrl,
      existsSync: (path) => existingPaths.has(path),
      fileURLToPath,
      runBuildSync,
    })

    expect(wasmPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(runBuildSync).toHaveBeenCalledTimes(1)
  })

  it('throws a clear error when the artifact is still missing after a build attempt', () => {
    const runBuildSync = vi.fn()

    expect(() =>
      ensureWasmBinaryPathForNode({
        importMetaUrl,
        existsSync: (path) => path === '/repo/packages/wasm-kernel/scripts/build.ts',
        fileURLToPath,
        runBuildSync,
      }),
    ).toThrow("Run 'pnpm wasm:build'")
    expect(runBuildSync).toHaveBeenCalledTimes(1)
  })
})
