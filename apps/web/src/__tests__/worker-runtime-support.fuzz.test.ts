import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { collectViewportCells, type ViewportCellPosition } from '../worker-runtime-support.js'

describe('worker runtime support fuzz', () => {
  it('should collect only cells inside invalidated row end bounds', async () => {
    await runProperty({
      suite: 'web/worker-runtime/viewport-row-invalidation-bounds',
      arbitrary: viewportAxisInvalidationArbitrary,
      predicate: async ({ viewport, startIndex, endIndex }) => {
        const actual = collectViewportCells(viewport, null, [], [{ startIndex, endIndex }], [])

        expect(cellKeys(actual)).toEqual(expectedRowInvalidationKeys(viewport, startIndex, endIndex))
      },
      parameters: { numRuns: 120 },
    })
  })

  it('should collect only cells inside invalidated column end bounds', async () => {
    await runProperty({
      suite: 'web/worker-runtime/viewport-column-invalidation-bounds',
      arbitrary: viewportAxisInvalidationArbitrary,
      predicate: async ({ viewport, startIndex, endIndex }) => {
        const actual = collectViewportCells(viewport, null, [], [], [{ startIndex, endIndex }])

        expect(cellKeys(actual)).toEqual(expectedColumnInvalidationKeys(viewport, startIndex, endIndex))
      },
      parameters: { numRuns: 120 },
    })
  })
})

const viewportAxisInvalidationArbitrary = fc
  .record({
    rowStart: fc.integer({ min: 0, max: 20 }),
    rowSpan: fc.integer({ min: 0, max: 8 }),
    colStart: fc.integer({ min: 0, max: 12 }),
    colSpan: fc.integer({ min: 0, max: 6 }),
    invalidationStartOffset: fc.integer({ min: -3, max: 5 }),
    invalidationSpan: fc.integer({ min: 0, max: 4 }),
  })
  .map(({ rowStart, rowSpan, colStart, colSpan, invalidationStartOffset, invalidationSpan }) => {
    const startIndex = rowStart + invalidationStartOffset
    return {
      viewport: {
        sheetName: 'Sheet1',
        rowStart,
        rowEnd: rowStart + rowSpan,
        colStart,
        colEnd: colStart + colSpan,
      },
      startIndex,
      endIndex: startIndex + invalidationSpan,
    }
  })

function cellKeys(cells: readonly ViewportCellPosition[]): string[] {
  return cells.map((cell) => `${cell.row}:${cell.col}`).toSorted()
}

function expectedRowInvalidationKeys(
  viewport: { readonly rowStart: number; readonly rowEnd: number; readonly colStart: number; readonly colEnd: number },
  startIndex: number,
  endIndex: number,
): string[] {
  const rowStart = Math.max(viewport.rowStart, startIndex)
  const rowEnd = Math.min(viewport.rowEnd, endIndex)
  if (rowStart > rowEnd) {
    return []
  }
  const keys: string[] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = viewport.colStart; col <= viewport.colEnd; col += 1) {
      keys.push(`${row}:${col}`)
    }
  }
  return keys.toSorted()
}

function expectedColumnInvalidationKeys(
  viewport: { readonly rowStart: number; readonly rowEnd: number; readonly colStart: number; readonly colEnd: number },
  startIndex: number,
  endIndex: number,
): string[] {
  const colStart = Math.max(viewport.colStart, startIndex)
  const colEnd = Math.min(viewport.colEnd, endIndex)
  if (colStart > colEnd) {
    return []
  }
  const keys: string[] = []
  for (let row = viewport.rowStart; row <= viewport.rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      keys.push(`${row}:${col}`)
    }
  }
  return keys.toSorted()
}
