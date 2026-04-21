import { describe, expect, test } from 'vitest'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { getGridMetrics } from '../gridMetrics.js'

describe('gridGeometry', () => {
  test('normalizes scroll into body world coordinates with frozen panes', () => {
    const metrics = getGridMetrics()
    const columns = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 })
    const rows = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 })
    const geometry = createGridGeometrySnapshotFromAxes({
      columns,
      dpr: 2,
      freezeCols: 2,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows,
      scrollLeft: 150,
      scrollTop: 40,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    expect(geometry.camera.frozenWidth).toBe(200)
    expect(geometry.camera.frozenHeight).toBe(20)
    expect(geometry.camera.bodyWorldX).toBe(350)
    expect(geometry.camera.bodyWorldY).toBe(60)
    expect(geometry.camera.bodyViewportWidth).toBe(274)
    expect(geometry.camera.bodyViewportHeight).toBe(176)
  })

  test('uses one transform model for body and frozen cells', () => {
    const metrics = getGridMetrics()
    const columns = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 })
    const rows = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 })
    const geometry = createGridGeometrySnapshotFromAxes({
      columns,
      dpr: 1,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows,
      scrollLeft: 50,
      scrollTop: 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    expect(geometry.cellScreenRect(0, 0)).toMatchObject({ height: 20, width: 100, x: metrics.rowMarkerWidth, y: metrics.headerHeight })
    expect(geometry.cellScreenRect(2, 3)).toMatchObject({
      height: 20,
      width: 100,
      x: metrics.rowMarkerWidth + 100 + (200 - 150),
      y: metrics.headerHeight + 20 + (60 - 30),
    })
    expect(geometry.columnHeaderScreenRect(2)).toMatchObject({ height: metrics.headerHeight, width: 100, x: 196, y: 0 })
    expect(geometry.rowHeaderScreenRect(3)).toMatchObject({ height: 20, width: metrics.rowMarkerWidth, x: 0, y: 74 })
  })

  test('hit-tests through body and frozen panes with hidden axes skipped', () => {
    const metrics = getGridMetrics()
    const columns = createGridAxisWorldIndex({
      axisLength: 20,
      defaultSize: 100,
      overrides: [{ hidden: true, index: 2 }],
    })
    const rows = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 })
    const geometry = createGridGeometrySnapshotFromAxes({
      columns,
      dpr: 1,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows,
      scrollLeft: 50,
      scrollTop: 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    expect(geometry.hitTestScreenPoint({ x: metrics.rowMarkerWidth + 10, y: metrics.headerHeight + 5 })).toEqual({ col: 0, row: 0 })
    expect(geometry.hitTestScreenPoint({ x: 205, y: 80 })).toEqual({ col: 3, row: 3 })
  })

  test('resolves range and fill-handle geometry through frozen/body panes', () => {
    const metrics = getGridMetrics()
    const columns = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 100 })
    const rows = createGridAxisWorldIndex({ axisLength: 20, defaultSize: 20 })
    const geometry = createGridGeometrySnapshotFromAxes({
      columns,
      dpr: 1,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics: metrics,
      hostHeight: 220,
      hostWidth: 520,
      rows,
      scrollLeft: 50,
      scrollTop: 10,
      sheetName: 'Sheet1',
      updatedAt: 100,
    })

    expect(geometry.rangeScreenRects({ x: 0, y: 0, width: 3, height: 3 })).toEqual([
      { height: 20, width: 100, x: 46, y: 24 },
      { height: 30, width: 100, x: 46, y: 44 },
      { height: 20, width: 150, x: 146, y: 24 },
      { height: 30, width: 150, x: 146, y: 44 },
    ])
    expect(geometry.fillHandleScreenRect({ x: 0, y: 0, width: 3, height: 3 })).toEqual({
      height: 8,
      width: 8,
      x: 292,
      y: 70,
    })
  })
})
