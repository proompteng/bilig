import { describe, expect, it } from 'vitest'
import { resolveBrowserTestPhases } from '../browser-test-phases.js'

describe('browser test phases', () => {
  it('keeps default browser tests focused on deterministic release coverage', () => {
    const phases = resolveBrowserTestPhases({ playwrightArgs: [], env: {} })

    expect(phases.map((phase) => phase.label)).toEqual([
      'parallel browser tests',
      'browser webgpu tests',
      'browser serial tests',
      'clipboard global tests',
    ])
    expect(phases[0]?.args).toEqual([
      '--workers=2',
      '--grep-invert',
      '@clipboard-global|@browser-serial|@browser-perf|@browser-deep|@browser-webgpu',
    ])
    expect(phases[1]).toEqual({
      label: 'browser webgpu tests',
      args: ['--workers=1', '--grep', '@browser-webgpu', '--grep-invert', '@browser-perf|@browser-deep'],
      env: { BILIG_BROWSER_WEBGPU: '1' },
    })
    expect(phases[2]?.args).toEqual(['--workers=1', '--grep', '@browser-serial', '--grep-invert', '@browser-webgpu'])
  })

  it('adds perf and deep only for the deep browser profile', () => {
    const phases = resolveBrowserTestPhases({
      playwrightArgs: [],
      env: {
        BILIG_BROWSER_INCLUDE_PERF: '1',
        BILIG_BROWSER_INCLUDE_DEEP: '1',
      },
    })

    expect(phases.map((phase) => phase.label)).toEqual([
      'parallel browser tests',
      'browser webgpu tests',
      'browser perf tests',
      'browser webgpu perf tests',
      'browser deep tests',
      'browser webgpu deep tests',
      'browser serial tests',
      'clipboard global tests',
    ])
    expect(phases.find((phase) => phase.label === 'browser perf tests')?.args).toEqual([
      '--workers=1',
      '--grep',
      '@browser-perf',
      '--grep-invert',
      '@browser-webgpu',
    ])
    expect(phases.find((phase) => phase.label === 'browser webgpu perf tests')).toEqual({
      label: 'browser webgpu perf tests',
      args: ['--workers=1', '--grep', '@browser-webgpu.*@browser-perf|@browser-perf.*@browser-webgpu'],
      env: { BILIG_BROWSER_WEBGPU: '1' },
    })
    expect(phases.find((phase) => phase.label === 'browser deep tests')?.args).toEqual([
      '--workers=1',
      '--grep',
      '@browser-deep',
      '--grep-invert',
      '@browser-webgpu',
      '--pass-with-no-tests',
    ])
    expect(phases.find((phase) => phase.label === 'browser webgpu deep tests')).toEqual({
      label: 'browser webgpu deep tests',
      args: ['--workers=1', '--grep', '@browser-webgpu.*@browser-deep|@browser-deep.*@browser-webgpu'],
      env: { BILIG_BROWSER_WEBGPU: '1' },
    })
    expect(phases.at(-1)?.label).toBe('clipboard global tests')
  })

  it('allows the default parallel browser worker cap to be configured', () => {
    const phases = resolveBrowserTestPhases({
      playwrightArgs: [],
      env: {
        BILIG_BROWSER_PARALLEL_WORKERS: '4',
      },
    })

    expect(phases[0]?.args).toEqual([
      '--workers=4',
      '--grep-invert',
      '@clipboard-global|@browser-serial|@browser-perf|@browser-deep|@browser-webgpu',
    ])
  })

  it('rejects malformed browser phase include flags instead of silently skipping coverage', () => {
    expect(() =>
      resolveBrowserTestPhases({
        playwrightArgs: [],
        env: {
          BILIG_BROWSER_INCLUDE_PERF: 'TRUE',
        },
      }),
    ).toThrow('BILIG_BROWSER_INCLUDE_PERF must be "1", "true", "0", or "false" when set, got TRUE')

    expect(() =>
      resolveBrowserTestPhases({
        playwrightArgs: [],
        env: {
          BILIG_BROWSER_CI_SMOKE: ' yes ',
        },
      }),
    ).toThrow('BILIG_BROWSER_CI_SMOKE must be "1", "true", "0", or "false" when set, got  yes ')
  })

  it('rejects malformed browser worker counts instead of silently using the default', () => {
    expect(() =>
      resolveBrowserTestPhases({
        playwrightArgs: [],
        env: {
          BILIG_BROWSER_PARALLEL_WORKERS: '4abc',
        },
      }),
    ).toThrow('BILIG_BROWSER_PARALLEL_WORKERS must be a positive integer, got 4abc')
  })

  it('keeps CI browser smoke explicit and small', () => {
    const phases = resolveBrowserTestPhases({
      playwrightArgs: [],
      env: {
        BILIG_BROWSER_CI_SMOKE: '1',
        BILIG_BROWSER_PARALLEL_WORKERS: '3',
      },
    })

    expect(phases).toEqual([
      {
        label: 'browser ci smoke tests',
        args: ['--workers=3', '--grep', '@browser-ci', '--grep-invert', '@browser-perf|@browser-deep|@browser-webgpu'],
      },
    ])
  })

  it('passes explicit Playwright arguments through unchanged', () => {
    expect(resolveBrowserTestPhases({ playwrightArgs: ['--grep', 'typegpu'], env: {} })).toEqual([
      {
        label: 'explicit browser tests',
        args: ['--grep', 'typegpu'],
      },
    ])
  })
})
