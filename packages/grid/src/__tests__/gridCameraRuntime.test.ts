import { describe, expect, it } from 'vitest'
import { unpackTileKey53 } from '../renderer-v3/tile-key.js'
import { GridAxisRuntime } from '../runtime/gridAxisRuntime.js'
import { GridCameraRuntime } from '../runtime/gridCameraRuntime.js'

const gridMetrics = {
  columnWidth: 100,
  headerHeight: 20,
  rowHeight: 10,
  rowMarkerWidth: 50,
}

describe('GridCameraRuntime', () => {
  it('computes visible regions from axis runtimes and scroll state', () => {
    const columns = new GridAxisRuntime({ axisLength: 1000, defaultSize: 100 })
    const rows = new GridAxisRuntime({ axisLength: 1000, defaultSize: 10 })
    const camera = new GridCameraRuntime({
      columns,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics,
      rows,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    expect(camera.snapshot()).toMatchObject({
      scrollLeft: 0,
      scrollTop: 0,
      seq: 1,
      visibleRegion: {
        freezeCols: 1,
        freezeRows: 1,
        range: {
          x: 1,
          y: 1,
        },
      },
    })

    const updated = camera.update({
      columns,
      dpr: 2,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics,
      rows,
      scrollLeft: 250,
      scrollTop: 35,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    expect(updated.seq).toBe(2)
    expect(updated.visibleRegion.range.x).toBe(3)
    expect(updated.visibleRegion.range.y).toBe(4)
    expect(updated.visibleRegion.tx).toBe(50)
    expect(updated.visibleRegion.ty).toBe(5)
  })

  it('returns fixed content tile keys for the current visible viewport', () => {
    const columns = new GridAxisRuntime({ axisLength: 1000, defaultSize: 100 })
    const rows = new GridAxisRuntime({ axisLength: 1000, defaultSize: 10 })
    const camera = new GridCameraRuntime({
      columns,
      gridMetrics,
      rows,
      scrollLeft: 250,
      scrollTop: 350,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    expect(camera.visibleTileKeys({ dprBucket: 2, sheetOrdinal: 5 }).map(unpackTileKey53)).toEqual([
      { colTile: 0, dprBucket: 2, rowTile: 1, sheetOrdinal: 5 },
    ])
  })
})
