import { describe, expect, it } from 'vitest'
import {
  ENGINE_COUNTER_KEYS,
  addEngineCounter,
  addEngineCounters,
  cloneEngineCounters,
  createEngineCounters,
  resetEngineCounters,
} from '../perf/engine-counters.js'

describe('engine counters', () => {
  it('initializes every planned performance counter to zero', () => {
    const counters = createEngineCounters()
    const zeroValues = Array.from({ length: ENGINE_COUNTER_KEYS.length }, () => 0)

    expect(Object.keys(counters).toSorted()).toEqual([...ENGINE_COUNTER_KEYS].toSorted())
    expect(Object.values(counters)).toEqual(zeroValues)
  })

  it('increments, clones, merges, and resets counters', () => {
    const counters = createEngineCounters()

    expect(addEngineCounter(counters, 'cellsRemapped')).toBe(1)
    expect(addEngineCounter(counters, 'cellsRemapped', 2)).toBe(3)

    const clone = cloneEngineCounters(counters)
    expect(clone).toEqual(counters)
    expect(clone).not.toBe(counters)

    addEngineCounters(counters, {
      formulasParsed: 4,
      exactIndexBuilds: 2,
      wasmFullUploads: 1,
    })

    expect(counters.cellsRemapped).toBe(3)
    expect(counters.formulasParsed).toBe(4)
    expect(counters.exactIndexBuilds).toBe(2)
    expect(counters.wasmFullUploads).toBe(1)

    resetEngineCounters(counters)
    expect(Object.values(counters)).toEqual(Array.from({ length: ENGINE_COUNTER_KEYS.length }, () => 0))
    expect(clone.cellsRemapped).toBe(3)
  })
})
