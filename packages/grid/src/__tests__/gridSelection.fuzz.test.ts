import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { clampSelectionRange, createRectangleSelectionFromRange, rectangleToAddresses, selectionToAddresses } from '../gridSelection.js'
import { resolveMovedRange, sameRectangle } from '../gridRangeMove.js'

describe('grid selection fuzz', () => {
  it('should keep rectangle selections and exported addresses aligned', async () => {
    await runProperty({
      suite: 'grid/selection/address-roundtrip',
      arbitrary: rectangleArbitrary,
      predicate: async (range) => {
        const selection = createRectangleSelectionFromRange(range)
        expect(selectionToAddresses(selection, 'A1')).toEqual(rectangleToAddresses(range))
        expect(selection.current.range).toEqual(range)
      },
    })
  })

  it('should preserve range size while clamping moved selections inside workbook bounds', async () => {
    await runProperty({
      suite: 'grid/selection/move-clamping',
      arbitrary: fc.record({
        sourceRange: boundedRectangleArbitrary,
        pointerCell: fc.tuple(fc.integer({ min: 0, max: 64 }), fc.integer({ min: 0, max: 64 })),
        anchorOffset: fc.tuple(fc.integer({ min: 0, max: 4 }), fc.integer({ min: 0, max: 4 })),
      }),
      predicate: async ({ sourceRange, pointerCell, anchorOffset }) => {
        const effectiveOffset = [
          Math.min(anchorOffset[0], sourceRange.width - 1),
          Math.min(anchorOffset[1], sourceRange.height - 1),
        ] as const
        const moved = resolveMovedRange(sourceRange, pointerCell, effectiveOffset)

        expect(moved.width).toBe(sourceRange.width)
        expect(moved.height).toBe(sourceRange.height)
        expect(moved.x).toBeGreaterThanOrEqual(0)
        expect(moved.y).toBeGreaterThanOrEqual(0)
        expect(moved.x + moved.width).toBeLessThanOrEqual(MAX_COLS)
        expect(moved.y + moved.height).toBeLessThanOrEqual(MAX_ROWS)
        expect(sameRectangle(moved, clampSelectionRange(moved))).toBe(true)
      },
    })
  })
})

// Helpers

const rectangleArbitrary = fc.record({
  x: fc.integer({ min: -8, max: 64 }),
  y: fc.integer({ min: -8, max: 64 }),
  width: fc.integer({ min: 1, max: 8 }),
  height: fc.integer({ min: 1, max: 8 }),
})

const boundedRectangleArbitrary = fc.record({
  x: fc.integer({ min: 0, max: 48 }),
  y: fc.integer({ min: 0, max: 48 }),
  width: fc.integer({ min: 1, max: 5 }),
  height: fc.integer({ min: 1, max: 5 }),
})
