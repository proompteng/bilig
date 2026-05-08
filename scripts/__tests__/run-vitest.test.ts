import { describe, expect, it } from 'vitest'

import { buildVitestArgs } from '../run-vitest.ts'

describe('run-vitest wrapper arguments', () => {
  it('bounds Vitest workers in CI by default', () => {
    expect(buildVitestArgs(['--run', 'sample.test.ts'], { BILIG_CI_PROFILE: 'fast' })).toEqual([
      '--run',
      'sample.test.ts',
      '--maxWorkers',
      '2',
    ])
  })

  it('preserves an explicit maxWorkers flag', () => {
    expect(buildVitestArgs(['--run', '--maxWorkers=1'], { BILIG_CI_PROFILE: 'fast' })).toEqual(['--run', '--maxWorkers=1'])
  })

  it('allows CI worker limit overrides', () => {
    expect(
      buildVitestArgs(['--run'], {
        BILIG_CI_PROFILE: 'fast',
        BILIG_VITEST_MAX_WORKERS: '3',
      }),
    ).toEqual(['--run', '--maxWorkers', '3'])
  })
})
