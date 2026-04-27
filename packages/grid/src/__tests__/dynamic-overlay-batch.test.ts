import { describe, expect, test, vi } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createColumnSliceSelection } from '../gridSelection.js'
import {
  DYNAMIC_OVERLAY_RECT_FLOAT_COUNT_V3,
  DYNAMIC_OVERLAY_RECT_INSTANCE_FLOAT_COUNT_V3,
  buildDynamicGridOverlayBatchV3,
} from '../renderer-v3/dynamic-overlay-batch.js'

describe('dynamic overlay batch v3', () => {
  test('builds selection, fill-handle, resize, and frozen separator rects from geometry without scene packets', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 50,
      scrollTop: 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      geometry,
      hoveredCell: [3, 3],
      resizeGuideColumn: 2,
      resizeGuideRow: 2,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      showFillHandle: true,
    })

    expect('packedScene' in overlay).toBe(false)
    expect(overlay.rectCount).toBe(overlay.fillRectCount + overlay.borderRectCount)
    expect(overlay.rects).toHaveLength(overlay.rectCount * DYNAMIC_OVERLAY_RECT_FLOAT_COUNT_V3)
    expect(overlay.rectInstances).toHaveLength(overlay.rectCount * DYNAMIC_OVERLAY_RECT_INSTANCE_FLOAT_COUNT_V3)
    expect(overlay.surfaceSize).toEqual({ height: 220, width: 520 })
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 146, y: 44, width: 150, height: 30 }),
        expect.objectContaining({ x: 290, y: 68, width: 12, height: 12 }),
        expect.objectContaining({ x: 297, y: 75, width: 98, height: 18 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 295, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 73, width: 520, height: 1 }),
        expect.objectContaining({ x: 290, y: 68, width: 1, height: 12 }),
        expect.objectContaining({ x: 145, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 43, width: 520, height: 1 }),
      ]),
    )
  })

  test('draws header and body axis selections from the live camera instead of resident panes', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 50,
      scrollTop: 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      activeHeaderDrag: { kind: 'column', index: 2 },
      geometry,
      gridSelection: createColumnSliceSelection(1, 3, 2),
      selectedCell: [1, 2],
      selectionRange: { x: 1, y: 2, width: 3, height: 1 },
      showFillHandle: false,
    })

    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 147, y: 1, width: 48, height: 22 }),
        expect.objectContaining({ x: 197, y: 1, width: 98, height: 22 }),
        expect.objectContaining({ x: 297, y: 1, width: 98, height: 22 }),
        expect.objectContaining({ x: 1, y: 55, width: 44, height: 18 }),
        expect.objectContaining({ x: 147, y: 25, width: 248, height: 18 }),
        expect.objectContaining({ x: 147, y: 45, width: 248, height: 174 }),
        expect.objectContaining({ x: 196, y: 21, width: 100, height: 3 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 146, y: 54, width: 50, height: 1 }),
        expect.objectContaining({ x: 96, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 395, y: 0, width: 1, height: 220 }),
      ]),
    )
  })

  test('builds visible header indexes without array sorting', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 200, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 2,
      freezeRows: 2,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 200, defaultSize: 20 }),
      scrollLeft: 150,
      scrollTop: 30,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })
    const toSortedSpy = vi.spyOn(Array.prototype, 'toSorted')

    buildDynamicGridOverlayBatchV3({
      geometry,
      gridSelection: createColumnSliceSelection(1, 4, 2),
      selectedCell: [1, 2],
      selectionRange: { x: 1, y: 2, width: 3, height: 1 },
      showFillHandle: false,
    })

    expect(toSortedSpy).not.toHaveBeenCalled()
    toSortedSpy.mockRestore()
  })
})

function readOverlayRects(batch: ReturnType<typeof buildDynamicGridOverlayBatchV3>): Array<{
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}> {
  const rects = []
  for (let index = 0; index < batch.rectCount; index += 1) {
    const offset = index * DYNAMIC_OVERLAY_RECT_FLOAT_COUNT_V3
    rects.push({
      x: batch.rects[offset + 0] ?? Number.NaN,
      y: batch.rects[offset + 1] ?? Number.NaN,
      width: batch.rects[offset + 2] ?? Number.NaN,
      height: batch.rects[offset + 3] ?? Number.NaN,
    })
  }
  return rects
}
