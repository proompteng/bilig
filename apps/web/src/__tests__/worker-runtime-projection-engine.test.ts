import { describe, expect, it, vi } from 'vitest'
import { acquireProjectionEngine, scheduleProjectionEngineMaterialization } from '../worker-runtime-projection-engine.js'

describe('worker runtime projection engine', () => {
  it('reuses an already installed engine without rebuilding', async () => {
    const engine = { name: 'installed' }
    const rebuildProjectionEngine = vi.fn()

    const result = await acquireProjectionEngine({
      getInstalledEngine: () => engine,
      getProjectionEnginePromise: () => null,
      getProjectionBuildVersion: () => 3,
      rebuildProjectionEngine,
      setProjectionOverlayScope: vi.fn(),
      installEngine: vi.fn(),
      setProjectionEnginePromise: vi.fn(),
      requireInstalledEngine: () => engine,
    })

    expect(result).toBe(engine)
    expect(rebuildProjectionEngine).not.toHaveBeenCalled()
  })

  it('reuses an in-flight projection build promise', async () => {
    const engine = { name: 'in-flight' }
    const promise = Promise.resolve(engine)

    const result = await acquireProjectionEngine({
      getInstalledEngine: () => null,
      getProjectionEnginePromise: () => promise,
      getProjectionBuildVersion: () => 9,
      rebuildProjectionEngine: vi.fn(),
      setProjectionOverlayScope: vi.fn(),
      installEngine: vi.fn(),
      setProjectionEnginePromise: vi.fn(),
      requireInstalledEngine: () => engine,
    })

    expect(result).toBe(engine)
  })

  it('installs a rebuilt projection engine when the build version stays current', async () => {
    const builtEngine = { name: 'built' }
    let projectionEnginePromise: Promise<typeof builtEngine> | null = null
    const setProjectionEnginePromise = vi.fn((promise: Promise<typeof builtEngine> | null) => {
      projectionEnginePromise = promise
    })
    const setProjectionOverlayScope = vi.fn()
    const installEngine = vi.fn()
    const requireInstalledEngine = vi.fn(() => builtEngine)
    let buildVersion = 4

    const result = await acquireProjectionEngine({
      getInstalledEngine: () => null,
      getProjectionEnginePromise: () => projectionEnginePromise,
      getProjectionBuildVersion: () => buildVersion,
      rebuildProjectionEngine: vi.fn(async () => ({
        engine: builtEngine,
        overlayScope: { kind: 'overlay' },
      })),
      setProjectionOverlayScope,
      installEngine,
      setProjectionEnginePromise,
      requireInstalledEngine,
    })

    expect(result).toBe(builtEngine)
    expect(setProjectionOverlayScope).toHaveBeenCalledWith({ kind: 'overlay' })
    expect(installEngine).toHaveBeenCalledWith(builtEngine)
    expect(requireInstalledEngine).toHaveBeenCalledTimes(1)
    expect(setProjectionEnginePromise).toHaveBeenCalledTimes(2)
    expect(setProjectionEnginePromise.mock.calls[0]?.[0]).toBeInstanceOf(Promise)
    expect(setProjectionEnginePromise.mock.calls[1]).toEqual([null])
    expect(buildVersion).toBe(4)
  })

  it('returns the current installed engine when a stale build resolves after a version bump', async () => {
    const staleEngine = { name: 'stale' }
    const installedEngine = { name: 'installed' }
    let buildVersion = 10
    let projectionEnginePromise: Promise<typeof staleEngine | typeof installedEngine> | null = null

    const result = await acquireProjectionEngine({
      getInstalledEngine: () => (buildVersion > 10 ? installedEngine : null),
      getProjectionEnginePromise: () => projectionEnginePromise,
      getProjectionBuildVersion: () => buildVersion,
      rebuildProjectionEngine: vi.fn(async () => {
        buildVersion = 11
        return { engine: staleEngine, overlayScope: { kind: 'stale' } }
      }),
      setProjectionOverlayScope: vi.fn(),
      installEngine: vi.fn(),
      setProjectionEnginePromise: vi.fn((promise: Promise<typeof staleEngine | typeof installedEngine> | null) => {
        projectionEnginePromise = promise
      }),
      requireInstalledEngine: () => installedEngine,
    })

    expect(result).toBe(installedEngine)
  })

  it('schedules projection materialization only when bootstrap is ready and no build exists', () => {
    vi.useFakeTimers()
    try {
      const getProjectionEngine = vi.fn(async () => undefined)

      scheduleProjectionEngineMaterialization({
        hasInstalledEngine: () => false,
        hasProjectionEnginePromise: () => false,
        hasBootstrapOptions: () => true,
        getProjectionBuildVersion: () => 2,
        getProjectionEngine,
        schedule: (callback) => {
          setTimeout(callback, 0)
        },
      })

      vi.runAllTimers()
      expect(getProjectionEngine).toHaveBeenCalledTimes(1)
    } finally {
      vi.useRealTimers()
    }
  })

  it('skips a scheduled projection materialization when the build version changes first', () => {
    vi.useFakeTimers()
    try {
      const getProjectionEngine = vi.fn(async () => undefined)
      let buildVersion = 5

      scheduleProjectionEngineMaterialization({
        hasInstalledEngine: () => false,
        hasProjectionEnginePromise: () => false,
        hasBootstrapOptions: () => true,
        getProjectionBuildVersion: () => buildVersion,
        getProjectionEngine,
        schedule: (callback) => {
          setTimeout(callback, 0)
        },
      })

      buildVersion = 6
      vi.runAllTimers()
      expect(getProjectionEngine).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
    }
  })
})
