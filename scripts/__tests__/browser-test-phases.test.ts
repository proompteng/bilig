import { describe, expect, it } from 'vitest'
import { resolveBrowserTestPhases } from '../browser-test-phases.js'

describe('browser test phases', () => {
  it('keeps default browser tests focused on deterministic release coverage', () => {
    const phases = resolveBrowserTestPhases({ playwrightArgs: [], env: {} })

    expect(phases.map((phase) => phase.label)).toEqual(['parallel browser tests', 'browser serial tests', 'clipboard global tests'])
    expect(phases[0]?.args).toEqual(['--grep-invert', '@clipboard-global|@browser-serial|@fuzz-browser|@browser-perf|@browser-deep'])
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
      'browser perf tests',
      'browser deep tests',
      'browser serial tests',
      'clipboard global tests',
      'browser fuzz tests',
    ])
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
