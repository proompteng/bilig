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
      '--grep-invert',
      '@clipboard-global|@browser-serial|@fuzz-browser|@browser-perf|@browser-deep|@browser-webgpu',
    ])
    expect(phases[1]).toEqual({
      label: 'browser webgpu tests',
      args: ['--workers=1', '--grep', '@browser-webgpu', '--grep-invert', '@browser-perf|@browser-deep'],
      env: { BILIG_BROWSER_WEBGPU: '1' },
    })
    expect(phases[2]?.args).toEqual(['--workers=1', '--grep', '@browser-serial', '--grep-invert', '@browser-webgpu'])
  })

  it('adds perf, deep, and fuzz only for the deep browser profile', () => {
    const phases = resolveBrowserTestPhases({
      playwrightArgs: [],
      env: {
        BILIG_BROWSER_INCLUDE_FUZZ: '1',
        BILIG_BROWSER_INCLUDE_PERF: '1',
        BILIG_BROWSER_INCLUDE_DEEP: '1',
        BILIG_FUZZ_PROFILE: 'nightly',
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
      'browser fuzz tests',
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
    ])
    expect(phases.find((phase) => phase.label === 'browser webgpu deep tests')).toEqual({
      label: 'browser webgpu deep tests',
      args: ['--workers=1', '--grep', '@browser-webgpu.*@browser-deep|@browser-deep.*@browser-webgpu'],
      env: { BILIG_BROWSER_WEBGPU: '1' },
    })
    expect(phases.at(-1)?.env).toEqual({
      BILIG_FUZZ_BROWSER: '1',
      BILIG_FUZZ_PROFILE: 'nightly',
      BILIG_FUZZ_CAPTURE: '1',
    })
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
