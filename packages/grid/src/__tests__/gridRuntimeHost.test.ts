import { describe, expect, it } from 'vitest'
import { unpackTileKey53 } from '../renderer-v3/tile-key.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'

const gridMetrics = {
  columnWidth: 100,
  headerHeight: 20,
  rowHeight: 10,
  rowMarkerWidth: 50,
}

describe('GridRuntimeHost', () => {
  it('composes camera, axis, tile-interest, and overlay runtime state', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    host.updateCamera({
      dpr: 2,
      gridMetrics,
      scrollLeft: 250,
      scrollTop: 350,
      viewportHeight: 60,
      viewportWidth: 250,
    })
    host.setOverlay('selection', { color: '#1a73e8', height: 20, kind: 'selection', width: 100, x: 0, y: 0 })

    expect(host.visibleTileKeys({ dprBucket: 2, sheetOrdinal: 9 }).map(unpackTileKey53)).toEqual([
      { colTile: 0, dprBucket: 2, rowTile: 1, sheetOrdinal: 9 },
    ])
    expect(host.buildOverlayBatch()).toMatchObject({
      cameraSeq: 2,
      count: 1,
    })
  })
})
