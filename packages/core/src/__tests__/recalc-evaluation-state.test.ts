import { describe, expect, it } from 'vitest'
import { utcDateToExcelSerial } from '@bilig/formula'
import { consumeVolatileRandomValues, createRecalcVolatileState, toOrderedUint32 } from '../engine/services/recalc-evaluation-state.js'

describe('recalc evaluation state helpers', () => {
  it('creates volatile state from the recalc clock', () => {
    const now = new Date('2026-05-16T12:34:56.000Z')
    const state = createRecalcVolatileState(() => now)

    expect(state).toEqual({
      nowSerial: utcDateToExcelSerial(now),
      randomValues: [],
      randomCursor: 0,
    })
  })

  it('consumes deterministic random values sequentially and extends lazily', () => {
    const values = [0.1, 0.2, 0.3]
    let cursor = 0
    const state = createRecalcVolatileState(() => new Date('2026-05-16T00:00:00.000Z'))
    const random = () => values[cursor++] ?? 0.9

    expect(Array.from(consumeVolatileRandomValues(state, 2, random))).toEqual([0.1, 0.2])
    expect(Array.from(consumeVolatileRandomValues(state, 1, random))).toEqual([0.3])
    expect(state.randomValues).toEqual([0.1, 0.2, 0.3])
    expect(state.randomCursor).toBe(3)
  })

  it('preserves typed ordered formula buffers and truncates arrays to the requested count', () => {
    const typed = Uint32Array.of(3, 2, 1)

    expect(toOrderedUint32(typed, 2)).toBe(typed)
    expect(Array.from(toOrderedUint32([7, 8, 9], 2))).toEqual([7, 8])
  })
})
