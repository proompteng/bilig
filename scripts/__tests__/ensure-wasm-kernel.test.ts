import { describe, expect, it, vi } from 'vitest'
import { ensureWasmKernelArtifact, resolveWasmKernelArtifactPath } from '../ensure-wasm-kernel.js'

describe('ensureWasmKernelArtifact', () => {
  it('returns the existing artifact path without spawning a build', () => {
    const spawnSync = vi.fn()
    const artifactPath = ensureWasmKernelArtifact({
      rootDir: '/repo',
      existsSync: (path) => path === '/repo/packages/wasm-kernel/build/release.wasm',
      spawnSync,
    })

    expect(artifactPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(spawnSync).not.toHaveBeenCalled()
  })

  it('builds the artifact when it is missing before startup', () => {
    const existingPaths = new Set<string>()
    const spawnSync = vi.fn((command: string, args: readonly string[], options: { cwd: string }) => {
      expect(command).toBe('pnpm')
      expect(args).toEqual(['wasm:build'])
      expect(options.cwd).toBe('/repo')
      existingPaths.add('/repo/packages/wasm-kernel/build/release.wasm')
      return {
        status: 0,
        stdout: Buffer.alloc(0),
        stderr: Buffer.alloc(0),
      }
    })

    const artifactPath = ensureWasmKernelArtifact({
      rootDir: '/repo',
      existsSync: (path) => existingPaths.has(path),
      spawnSync,
      env: {},
    })

    expect(artifactPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(spawnSync).toHaveBeenCalledTimes(1)
  })

  it('throws a clear error when the build fails', () => {
    expect(() =>
      ensureWasmKernelArtifact({
        rootDir: '/repo',
        existsSync: () => false,
        spawnSync: () => ({
          status: 1,
          stdout: Buffer.from(''),
          stderr: Buffer.from('boom'),
        }),
        env: {},
      }),
    ).toThrow('Failed to build wasm kernel artifact: boom')
  })

  it('throws when the build exits successfully without producing the artifact', () => {
    expect(() =>
      ensureWasmKernelArtifact({
        rootDir: '/repo',
        existsSync: () => false,
        spawnSync: () => ({
          status: 0,
          stdout: Buffer.alloc(0),
          stderr: Buffer.alloc(0),
        }),
        env: {},
      }),
    ).toThrow("did not create '/repo/packages/wasm-kernel/build/release.wasm'")
  })
})

describe('resolveWasmKernelArtifactPath', () => {
  it('derives the release.wasm path from the workspace root', () => {
    expect(resolveWasmKernelArtifactPath('/repo')).toBe('/repo/packages/wasm-kernel/build/release.wasm')
  })
})
