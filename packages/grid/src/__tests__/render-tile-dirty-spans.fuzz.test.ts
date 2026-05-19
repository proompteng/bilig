import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import type { GridRenderTile, GridRenderTileDirtySpan } from '../renderer-v3/render-tile-source.js'
import { resolveGridRenderTileDirtySpansV3 } from '../renderer-v3/render-tile-dirty-spans.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'

describe('render tile dirty spans fuzz', () => {
  it('should emit sorted bounded non-overlapping dirty spans for generated tile-local damage', async () => {
    await runProperty({
      suite: 'grid/renderer-v3/render-tile-dirty-spans/bounded-merged-spans',
      arbitrary: dirtyTileArbitrary,
      predicate: async (tile) => {
        const dirty = resolveGridRenderTileDirtySpansV3(tile)

        expectSortedBoundedSpans(dirty.rectSpans, tile.rectCount)
        expectSortedBoundedSpans(dirty.textSpans, tile.textCount)
        expectSortedBoundedSpans(dirty.glyphSpans, tile.textCount)
      },
      parameters: { numRuns: 120 },
    })
  })
})

const dirtyTileArbitrary: fc.Arbitrary<GridRenderTile> = fc
  .record({
    rowCount: fc.integer({ min: 1, max: 8 }),
    colCount: fc.integer({ min: 1, max: 8 }),
    dirtyRanges: fc.array(
      fc.record({
        rowStart: fc.integer({ min: 0, max: 7 }),
        rowEnd: fc.integer({ min: 0, max: 7 }),
        colStart: fc.integer({ min: 0, max: 7 }),
        colEnd: fc.integer({ min: 0, max: 7 }),
        mask: fc.constantFrom(
          DirtyMaskV3.Value | DirtyMaskV3.Text,
          DirtyMaskV3.Style | DirtyMaskV3.Rect,
          DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect,
          DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect,
        ),
      }),
      { minLength: 0, maxLength: 5 },
    ),
  })
  .map(({ rowCount, colCount, dirtyRanges }) => {
    const normalizedRanges = dirtyRanges.map((range) => {
      const rowStart = Math.min(range.rowStart, range.rowEnd, rowCount - 1)
      const rowEnd = Math.min(Math.max(range.rowStart, range.rowEnd), rowCount - 1)
      const colStart = Math.min(range.colStart, range.colEnd, colCount - 1)
      const colEnd = Math.min(Math.max(range.colStart, range.colEnd), colCount - 1)
      return { ...range, rowStart, rowEnd, colStart, colEnd }
    })
    return createTile({
      bounds: { rowStart: 0, rowEnd: rowCount - 1, colStart: 0, colEnd: colCount - 1 },
      dirtyLocalRows: Uint32Array.from(normalizedRanges.flatMap((range) => [range.rowStart, range.rowEnd])),
      dirtyLocalCols: Uint32Array.from(normalizedRanges.flatMap((range) => [range.colStart, range.colEnd])),
      dirtyMasks: Uint32Array.from(normalizedRanges.map((range) => range.mask)),
      rectCount: rowCount * colCount,
      rectInstances: new Float32Array(rowCount * colCount * 20),
      textCount: rowCount * colCount,
      textMetrics: new Float32Array(rowCount * colCount * 8),
      textRuns: createTextRuns(rowCount, colCount),
    })
  })

function expectSortedBoundedSpans(spans: readonly GridRenderTileDirtySpan[], count: number): void {
  let previousEnd = 0
  for (const span of spans) {
    expect(span.offset).toBeGreaterThanOrEqual(previousEnd)
    expect(span.length).toBeGreaterThan(0)
    expect(span.offset + span.length).toBeLessThanOrEqual(count)
    previousEnd = span.offset + span.length
  }
}

function createTile(overrides: Partial<GridRenderTile>): GridRenderTile {
  return {
    bounds: { colEnd: 0, colStart: 0, rowEnd: 0, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 1,
    rectInstances: new Float32Array(20),
    textCount: 1,
    textMetrics: new Float32Array(8),
    textRuns: [],
    tileId: 101,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
    ...overrides,
  }
}

function createTextRuns(rowCount: number, colCount: number): GridRenderTile['textRuns'] {
  const runs: GridRenderTile['textRuns'] = []
  for (let row = 0; row < rowCount; row += 1) {
    for (let col = 0; col < colCount; col += 1) {
      runs.push({
        align: 'left',
        clipHeight: 22,
        clipWidth: 104,
        clipX: 0,
        clipY: 0,
        color: '#111111',
        col,
        font: '400 11px sans-serif',
        fontSize: 11,
        height: 22,
        row,
        strike: false,
        text: `${row}:${col}`,
        underline: false,
        width: 104,
        wrap: false,
        x: 0,
        y: 0,
      })
    }
  }
  return runs
}
