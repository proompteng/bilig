import { describe, expect, test, vi } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'

function createGeometry(seq: number) {
  const metrics = getGridMetrics()
  return createGridGeometrySnapshotFromAxes({
    columns: createGridAxisWorldIndex({ axisLength: 10, defaultSize: metrics.columnWidth }),
    dpr: 1,
    gridMetrics: metrics,
    hostHeight: 240,
    hostWidth: 480,
    rows: createGridAxisWorldIndex({ axisLength: 10, defaultSize: metrics.rowHeight }),
    scrollLeft: seq * 10,
    scrollTop: seq * 5,
    seq,
    sheetName: 'Sheet1',
    updatedAt: seq,
  })
}

describe('GridCameraStore', () => {
  test('publishes distinct camera sequences once', () => {
    const store = new GridCameraStore()
    const listener = vi.fn()
    store.subscribe(listener)
    const first = createGeometry(1)
    const duplicateSeq = createGeometry(1)
    const second = createGeometry(2)

    store.setSnapshot(first)
    store.setSnapshot(duplicateSeq)
    store.setSnapshot(second)

    expect(listener).toHaveBeenCalledTimes(2)
    expect(store.getSnapshot()).toBe(second)
  })
})
