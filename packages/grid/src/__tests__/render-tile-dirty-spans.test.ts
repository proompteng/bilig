import { describe, expect, test } from 'vitest'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { resolveFullGridRenderTileDirtySpansV3, resolveGridRenderTileDirtySpansV3 } from '../renderer-v3/render-tile-dirty-spans.js'
import { DirtyMaskV3 } from '../renderer-v3/tile-damage-index.js'

function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  return {
    bounds: { colEnd: 3, colStart: 0, rowEnd: 3, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 16,
    rectInstances: new Float32Array(16 * 20),
    textCount: 2,
    textMetrics: new Float32Array(2 * 8),
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

describe('render tile dirty spans v3', () => {
  test('uses full dirty spans when tile-local metadata is absent', () => {
    expect(resolveGridRenderTileDirtySpansV3(createTile())).toEqual(resolveFullGridRenderTileDirtySpansV3(createTile()))
  })

  test('omits rect spans for text-only local updates', () => {
    const spans = resolveGridRenderTileDirtySpansV3(
      createTile({
        dirtyLocalCols: new Uint32Array([1, 1]),
        dirtyLocalRows: new Uint32Array([1, 1]),
        dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
      }),
    )

    expect(spans.rectSpans).toEqual([])
    expect(spans.textSpans).toEqual([{ offset: 0, length: 2 }])
  })

  test('maps text damage to tile-local text-run spans when runs carry cell coordinates', () => {
    const spans = resolveGridRenderTileDirtySpansV3(
      createTile({
        dirtyLocalCols: new Uint32Array([1, 1]),
        dirtyLocalRows: new Uint32Array([2, 2]),
        dirtyMasks: new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]),
        textCount: 3,
        textRuns: [createTextRun({ col: 0, row: 2 }), createTextRun({ col: 1, row: 2 }), createTextRun({ col: 1, row: 3 })],
      }),
    )

    expect(spans.rectSpans).toEqual([])
    expect(spans.textSpans).toEqual([{ offset: 1, length: 1 }])
  })

  test('maps exact cell-local rect damage to row-major rect instance spans', () => {
    const spans = resolveGridRenderTileDirtySpansV3(
      createTile({
        dirtyLocalCols: new Uint32Array([1, 2]),
        dirtyLocalRows: new Uint32Array([2, 2]),
        dirtyMasks: new Uint32Array([DirtyMaskV3.Style | DirtyMaskV3.Rect]),
      }),
    )

    expect(spans.rectSpans).toEqual([{ offset: 9, length: 2 }])
    expect(spans.textSpans).toEqual([{ offset: 0, length: 2 }])
  })

  test('merges unsorted overlapping rect dirty spans without a sorted copy', () => {
    const spans = resolveGridRenderTileDirtySpansV3(
      createTile({
        dirtyLocalCols: new Uint32Array([1, 3, 0, 1]),
        dirtyLocalRows: new Uint32Array([2, 2, 2, 2]),
        dirtyMasks: new Uint32Array([DirtyMaskV3.Style | DirtyMaskV3.Rect, DirtyMaskV3.Style | DirtyMaskV3.Rect]),
      }),
    )

    expect(spans.rectSpans).toEqual([{ offset: 8, length: 4 }])
  })

  test('falls back to full rect damage when rect instances are not one-per-cell', () => {
    const spans = resolveGridRenderTileDirtySpansV3(
      createTile({
        dirtyLocalCols: new Uint32Array([1, 1]),
        dirtyLocalRows: new Uint32Array([1, 1]),
        dirtyMasks: new Uint32Array([DirtyMaskV3.Style | DirtyMaskV3.Rect]),
        rectCount: 20,
      }),
    )

    expect(spans.rectSpans).toEqual([{ offset: 0, length: 20 }])
  })
})

function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  return {
    align: 'left',
    clipHeight: 22,
    clipWidth: 104,
    clipX: 0,
    clipY: 0,
    color: '#111111',
    font: '400 11px sans-serif',
    fontSize: 11,
    height: 22,
    strike: false,
    text: 'A1',
    underline: false,
    width: 104,
    wrap: false,
    x: 0,
    y: 0,
    ...overrides,
  }
}
