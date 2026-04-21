import { describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { buildDynamicGridOverlayPacket } from '../renderer-v2/dynamic-overlay-packet.js'

describe('dynamic overlay packet', () => {
  test('builds selection, fill-handle, resize, and frozen separator rects from geometry', () => {
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

    const overlay = buildDynamicGridOverlayPacket({
      geometry,
      hoveredCell: [3, 3],
      resizeGuideColumn: 2,
      resizeGuideRow: 2,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      showFillHandle: true,
    })

    expect(overlay.textScene.items).toEqual([])
    expect(overlay.gpuScene.fillRects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 146, y: 44, width: 150, height: 30 }),
        expect.objectContaining({ x: 292, y: 70, width: 8, height: 8 }),
        expect.objectContaining({ x: 297, y: 75, width: 98, height: 18 }),
      ]),
    )
    expect(overlay.gpuScene.borderRects).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ x: 295, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 73, width: 520, height: 1 }),
        expect.objectContaining({ x: 145, y: 0, width: 1, height: 220 }),
        expect.objectContaining({ x: 0, y: 43, width: 520, height: 1 }),
      ]),
    )
  })
})
