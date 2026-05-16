import { describe, expect, it } from 'vitest'
import {
  DEFAULT_ITERATION_COUNT,
  DEFAULT_ITERATION_DELTA,
  resolveRecalcIterationSettings,
} from '../engine/services/recalc-iteration-settings.js'

describe('resolveRecalcIterationSettings', () => {
  it('uses explicit positive iteration settings', () => {
    expect(
      resolveRecalcIterationSettings({
        mode: 'automatic',
        iterate: true,
        iterateCount: 32,
        iterateDelta: '0.0001',
      }),
    ).toEqual({
      enabled: true,
      count: 32,
      delta: 0.0001,
    })
  })

  it('falls back for invalid iteration count and malformed delta text', () => {
    expect(
      resolveRecalcIterationSettings({
        mode: 'automatic',
        iterate: true,
        iterateCount: 0,
        iterateDelta: '',
      }),
    ).toEqual({
      enabled: true,
      count: DEFAULT_ITERATION_COUNT,
      delta: DEFAULT_ITERATION_DELTA,
    })

    expect(
      resolveRecalcIterationSettings({
        mode: 'automatic',
        iterate: true,
        iterateCount: 3.5,
        iterateDelta: ' 0.01 ',
      }),
    ).toEqual({
      enabled: true,
      count: DEFAULT_ITERATION_COUNT,
      delta: DEFAULT_ITERATION_DELTA,
    })
  })

  it('keeps iteration disabled while preserving sanitized defaults', () => {
    expect(
      resolveRecalcIterationSettings({
        mode: 'automatic',
        iterate: false,
        iterateCount: null,
        iterateDelta: '-0.1',
      }),
    ).toEqual({
      enabled: false,
      count: DEFAULT_ITERATION_COUNT,
      delta: DEFAULT_ITERATION_DELTA,
    })
  })
})
