import { describe, expect, test, vi } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createColumnSliceSelection, createGridSelection, createRangeSelection, createRowSliceSelection } from '../gridSelection.js'
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
      gridSelection: createRangeSelection(createGridSelection(1, 1), [1, 1], [2, 2]),
      hoveredCell: [3, 3],
      resizeGuideColumn: 2,
      resizeGuideRow: 2,
      selectedCell: [1, 1],
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
        expect.objectContaining({ x: 146, y: 44, width: 150, height: 1 }),
        expect.objectContaining({ x: 146, y: 44, width: 1, height: 30 }),
        expect.objectContaining({ x: 292.5, y: 70.5, width: 7, height: 7 }),
        expect.objectContaining({ x: 297, y: 75, width: 98, height: 18 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 197, y: 45, width: 98, height: 8 }),
        expect.objectContaining({ x: 147, y: 55, width: 148, height: 18 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 295, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 73, width: 520, height: 1 }),
        expect.objectContaining({ x: 145, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 43, width: 520, height: 1 }),
      ]),
    )
    expect(readOverlayRects(overlay)).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ x: 291.5, y: 70.5, width: 1, height: 7 })]),
    )
  })

  test('draws axis selection headers and guides from the live camera without masking cell fills', () => {
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
        expect.objectContaining({ x: 196, y: 21, width: 100, height: 3 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 147, y: 25, width: 248, height: 18 }),
        expect.objectContaining({ x: 147, y: 45, width: 248, height: 8 }),
        expect.objectContaining({ x: 197, y: 55, width: 198, height: 18 }),
        expect.objectContaining({ x: 147, y: 75, width: 248, height: 144 }),
      ]),
    )
    expect(readOverlayRects(overlay)).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ x: 147, y: 55, width: 48, height: 18 })]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 146, y: 54, width: 50, height: 2 }),
        expect.objectContaining({ x: 96, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 395, y: 0, width: 1, height: 220 }),
      ]),
    )
  })

  test('draws a complete TypeGPU active-cell border inside the selected range', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 0,
      freezeRows: 0,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      geometry,
      gridSelection: createRangeSelection(createGridSelection(3, 3), [3, 3], [1, 1]),
      selectedCell: [3, 3],
      selectionRange: { x: 1, y: 1, width: 3, height: 3 },
      showFillHandle: true,
    })

    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 346, y: 84, width: 100, height: 2 }),
        expect.objectContaining({ x: 346, y: 84, width: 2, height: 20 }),
        expect.objectContaining({ x: 346, y: 102, width: 100, height: 2 }),
        expect.objectContaining({ x: 444, y: 84, width: 2, height: 20 }),
      ]),
    )
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 147, y: 45, width: 298, height: 38 }),
        expect.objectContaining({ x: 147, y: 85, width: 198, height: 18 }),
      ]),
    )
  })

  test('draws resize preview guides from overlay dimensions without mutating axis geometry', () => {
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
      resizeGuideColumn: 2,
      resizeGuideColumnWidth: 140,
      resizeGuideRow: 2,
      resizeGuideRowHeight: 35,
      selectionRange: null,
      showFillHandle: false,
    })

    expect(geometry.columns.sizeOf(2)).toBe(100)
    expect(geometry.rows.sizeOf(2)).toBe(20)
    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 335, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 88, width: 520, height: 1 }),
      ]),
    )
  })

  test('can exclude hover from TypeGPU overlay batches so normal selection stays allocation-free', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 0,
      freezeRows: 0,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      geometry,
      hoveredCell: [3, 3],
      selectionRange: null,
      showFillHandle: false,
      showHoverOverlay: false,
      showSelectionOverlay: false,
    })

    expect(overlay.rectCount).toBe(0)
    expect(overlay.rectInstances).toHaveLength(DYNAMIC_OVERLAY_RECT_INSTANCE_FLOAT_COUNT_V3)
  })

  test('draws row-selection body highlights through the dynamic overlay', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 0,
      freezeRows: 0,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      geometry,
      gridSelection: createRowSliceSelection(1, 2, 4),
      selectedCell: [1, 2],
      selectionRange: { x: 1, y: 2, width: 1, height: 3 },
      showFillHandle: false,
    })

    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 47, y: 65, width: 98, height: 18 }),
        expect.objectContaining({ x: 247, y: 65, width: 272, height: 18 }),
        expect.objectContaining({ x: 47, y: 85, width: 472, height: 38 }),
        expect.objectContaining({ x: 146, y: 64, width: 100, height: 2 }),
      ]),
    )
    expect(readOverlayRects(overlay)).not.toEqual(
      expect.arrayContaining([expect.objectContaining({ x: 147, y: 65, width: 98, height: 18 })]),
    )
  })

  test('draws fill and review preview ranges through the V3 overlay batch', () => {
    const metrics = getGridMetrics()
    const geometry = createGridGeometrySnapshotFromAxes({
      columns: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 }),
      dpr: 2,
      freezeCols: 0,
      freezeRows: 0,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows: createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 }),
      scrollLeft: 0,
      scrollTop: 0,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    const overlay = buildDynamicGridOverlayBatchV3({
      fillPreviewRange: { x: 2, y: 2, width: 2, height: 1 },
      geometry,
      previewRects: [
        {
          role: 'target',
          bounds: { x: 146, y: 64, width: 200, height: 20 },
        },
      ],
      selectionRange: null,
      showFillHandle: false,
    })

    expect(readOverlayRects(overlay)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 246, y: 64, width: 200, height: 1 }),
        expect.objectContaining({ x: 146, y: 64, width: 200, height: 20 }),
        expect.objectContaining({ x: 146, y: 64, width: 200, height: 1 }),
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
