import { describe, expect, it } from 'vitest'
import { intersectRangeBounds, normalizeRange } from '../engine-range-utils.js'

describe('engine range utils', () => {
  it('normalizes reversed addresses into ascending bounds', () => {
    expect(
      normalizeRange({
        sheetName: 'Sheet1',
        startAddress: 'D5',
        endAddress: 'B2',
      }),
    ).toEqual({
      startRow: 1,
      endRow: 4,
      startCol: 1,
      endCol: 3,
    })
  })

  it('intersects metadata bounds with the requested range', () => {
    expect(
      intersectRangeBounds(
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'E6',
        },
        {
          startRow: 4,
          endRow: 8,
          startCol: 1,
          endCol: 3,
        },
      ),
    ).toEqual({
      startRow: 4,
      endRow: 5,
      startCol: 1,
      endCol: 3,
    })
  })

  it('returns undefined when ranges do not overlap', () => {
    expect(
      intersectRangeBounds(
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        {
          startRow: 5,
          endRow: 7,
          startCol: 5,
          endCol: 6,
        },
      ),
    ).toBeUndefined()
  })
})
