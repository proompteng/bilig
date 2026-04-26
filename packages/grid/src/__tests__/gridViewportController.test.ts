import { describe, expect, it } from 'vitest'
import {
  resolveGridRenderScrollTransform,
  sameViewportBounds,
  sameVisibleRegionWindow,
  tileKeyToViewport,
} from '../gridViewportController.js'

describe('gridViewportController', () => {
  it('compares only the resident visible window fields that trigger React state', () => {
    const base = {
      freezeCols: 1,
      freezeRows: 2,
      range: { x: 4, y: 6, width: 12, height: 20 },
      tx: 0,
      ty: 0,
    }

    expect(
      sameVisibleRegionWindow(base, {
        ...base,
        tx: 19,
        ty: 23,
      }),
    ).toBe(true)
    expect(
      sameVisibleRegionWindow(base, {
        ...base,
        range: { ...base.range, x: 5 },
      }),
    ).toBe(false)
  })

  it('compares viewport bounds and converts tile keys without changing shape', () => {
    const viewport = { rowStart: 10, rowEnd: 20, colStart: 30, colEnd: 40 }

    expect(sameViewportBounds(viewport, { ...viewport })).toBe(true)
    expect(sameViewportBounds(viewport, { ...viewport, colEnd: 41 })).toBe(false)
    expect(tileKeyToViewport(viewport)).toEqual(viewport)
  })

  it('resolves render transforms against the resident render viewport', () => {
    expect(
      resolveGridRenderScrollTransform({
        nextVisibleRegion: {
          freezeCols: 0,
          freezeRows: 0,
          range: { x: 5, y: 6, width: 12, height: 20 },
          tx: 4,
          ty: 3,
        },
        renderViewport: { rowStart: 4, rowEnd: 35, colStart: 3, colEnd: 130 },
        sortedColumnWidthOverrides: [
          [3, 120],
          [4, 80],
        ],
        sortedRowHeightOverrides: [[4, 30]],
        defaultColumnWidth: 100,
        defaultRowHeight: 20,
      }),
    ).toEqual({ renderTx: 204, renderTy: 53 })
  })
})
