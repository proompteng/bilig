import { describe, expect, it } from 'vitest'
import { GridAxisRuntime } from '../runtime/gridAxisRuntime.js'

describe('GridAxisRuntime', () => {
  it('answers offset and span queries from an update-owned axis index', () => {
    const axis = new GridAxisRuntime({
      axisLength: 10,
      defaultSize: 20,
      overrides: [{ index: 2, size: 40 }],
      seq: 3,
    })

    expect(axis.snapshot()).toEqual({ axisLength: 10, defaultSize: 20, seq: 3 })
    expect(axis.sizeAt(2)).toBe(40)
    expect(axis.offsetOf(3)).toBe(80)
    expect(axis.span(1, 4)).toBe(80)
    expect(axis.tileOrigin(3)).toBe(80)
  })

  it('updates revisions and visible ranges without rebuilding on every query', () => {
    const axis = new GridAxisRuntime({ axisLength: 100, defaultSize: 10 })

    expect(axis.visibleRangeForOffset(35, 40)).toEqual({ start: 3, endExclusive: 7 })

    axis.update({
      seq: 9,
      overrides: [
        { index: 3, size: 20 },
        { hidden: true, index: 4, size: 10 },
      ],
    })

    expect(axis.snapshot().seq).toBe(9)
    expect(axis.offsetOf(5)).toBe(50)
    expect(axis.visibleRangeForOffset(35, 40)).toEqual({ start: 3, endExclusive: 7 })
  })
})
