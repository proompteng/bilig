import { describe, expect, test } from 'vitest'
import { createGridCameraSnapshot } from '../gridCamera.js'
import { getGridMetrics } from '../gridMetrics.js'

describe('gridCamera', () => {
  test('creates a camera snapshot from native scroll input', () => {
    const gridMetrics = getGridMetrics()
    const camera = createGridCameraSnapshot({
      columnWidths: {},
      dpr: 2,
      gridMetrics,
      rowHeights: {},
      scrollLeft: gridMetrics.columnWidth * 3 + 12,
      scrollTop: gridMetrics.rowHeight * 5 + 7,
      updatedAt: 100,
      viewportHeight: 180,
      viewportWidth: 480,
    })

    expect(camera.tx).toBe(12)
    expect(camera.ty).toBe(7)
    expect(camera.visibleViewport).toMatchObject({
      colStart: 3,
      rowStart: 5,
    })
    expect(camera.residentViewport).toMatchObject({
      colStart: 0,
      rowStart: 0,
    })
  })

  test('records velocity from the previous camera snapshot', () => {
    const gridMetrics = getGridMetrics()
    const previous = createGridCameraSnapshot({
      columnWidths: {},
      dpr: 1,
      gridMetrics,
      rowHeights: {},
      scrollLeft: 100,
      scrollTop: 200,
      updatedAt: 100,
      viewportHeight: 180,
      viewportWidth: 480,
    })
    const next = createGridCameraSnapshot({
      columnWidths: {},
      dpr: 1,
      gridMetrics,
      previous,
      rowHeights: {},
      scrollLeft: 180,
      scrollTop: 260,
      updatedAt: 116,
      viewportHeight: 180,
      viewportWidth: 480,
    })

    expect(next.velocityX).toBe(5_000)
    expect(next.velocityY).toBe(3_750)
  })
})
