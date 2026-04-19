import { describe, expect, it } from 'vitest'
import { buildVitestFuzzCommand, resolveVitestFuzzMaxWorkers } from '../run-fuzz-config.js'

describe('run fuzz config', () => {
  it('caps vitest fuzz workers to a conservative subset of host parallelism', () => {
    expect(resolveVitestFuzzMaxWorkers(1)).toBe(1)
    expect(resolveVitestFuzzMaxWorkers(2)).toBe(1)
    expect(resolveVitestFuzzMaxWorkers(3)).toBe(2)
    expect(resolveVitestFuzzMaxWorkers(8)).toBe(4)
    expect(resolveVitestFuzzMaxWorkers(32)).toBe(4)
  })

  it('appends the maxWorkers flag to the vitest fuzz command', () => {
    expect(buildVitestFuzzCommand(['packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts'], 8)).toEqual([
      'pnpm',
      'exec',
      'vitest',
      'run',
      'packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts',
      '--maxWorkers',
      '4',
    ])
  })
})
