import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { Rectangle } from '../gridTypes.js'
import { unpackTileKey53 } from '../renderer-v3/tile-key.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'

const gridMetrics = {
  columnWidth: 100,
  headerHeight: 20,
  rowHeight: 10,
  rowMarkerWidth: 50,
}

function createEmptyCellSnapshot(address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value: { tag: ValueTag.Empty },
    version: 0,
  }
}

const LOCAL_EMPTY_ENGINE: GridEngineLike = {
  getCell: (_sheetName, address) => createEmptyCellSnapshot(address),
  getCellStyle: () => undefined,
  subscribeCells: () => () => {},
  workbook: {
    getSheet: () => undefined,
  },
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
    expect(host.buildTileInterest({ dprBucket: 2, reason: 'scroll', sheetId: 4, sheetOrdinal: 9 })).toMatchObject({
      seq: 1,
      sheetId: 4,
      sheetOrdinal: 9,
      cameraSeq: 2,
      axisSeqX: 0,
      axisSeqY: 0,
      freezeSeq: 1,
      visibleTileKeys: host.visibleTileKeys({ dprBucket: 2, sheetOrdinal: 9 }),
      reason: 'scroll',
    })
  })

  it('keeps freeze state in the runtime transaction when update input omits it', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    expect(host.snapshot().freezeSeq).toBe(1)
    expect(host.snapshot().camera.visibleRegion.freezeCols).toBe(1)
    host.updateCamera({
      dpr: 1,
      gridMetrics,
      scrollLeft: 100,
      scrollTop: 100,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    expect(host.snapshot().freezeSeq).toBe(1)
    expect(host.snapshot().camera.visibleRegion.freezeCols).toBe(1)

    host.updateCamera({
      dpr: 1,
      freezeCols: 2,
      freezeRows: 1,
      gridMetrics,
      scrollLeft: 100,
      scrollTop: 100,
      viewportHeight: 80,
      viewportWidth: 300,
    })
    expect(host.snapshot().freezeSeq).toBe(2)
    expect(host.snapshot().camera.visibleRegion.freezeCols).toBe(2)
  })

  it('updates axis revisions and tile origins independently from camera ownership', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    host.updateAxes({
      columnSeq: 41,
      columns: [{ index: 0, size: 160 }],
      rowSeq: 42,
      rows: [{ index: 1, size: 30 }],
    })
    host.updateCamera({
      dpr: 1,
      gridMetrics,
      scrollLeft: 0,
      scrollTop: 0,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    expect(host.snapshot()).toMatchObject({
      axisSeqX: 41,
      axisSeqY: 42,
    })
    expect(host.columns.tileOrigin(1)).toBe(160)
    expect(host.rows.tileOrigin(2)).toBe(40)
    expect(host.buildTileInterest({ dprBucket: 1, reason: 'scroll', sheetId: 4, sheetOrdinal: 9 })).toMatchObject({
      axisSeqX: 41,
      axisSeqY: 42,
    })
  })

  it('owns resident viewport and header interest computation outside the render hook', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      freezeCols: 2,
      freezeRows: 1,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    const first = host.resolveViewportResidency({
      freezeCols: 2,
      freezeRows: 1,
      sceneRevision: 0,
      visibleRegion: {
        freezeCols: 2,
        freezeRows: 1,
        range: { height: 8, width: 10, x: 260, y: 110 },
        tx: 0,
        ty: 0,
      },
    })
    const sameWindow = host.resolveViewportResidency({
      freezeCols: 2,
      freezeRows: 1,
      sceneRevision: 1,
      visibleRegion: {
        freezeCols: 2,
        freezeRows: 1,
        range: { height: 8, width: 10, x: 261, y: 111 },
        tx: 0,
        ty: 0,
      },
    })

    expect(first.residentViewport).toEqual({
      colEnd: 511,
      colStart: 256,
      rowEnd: 191,
      rowStart: 96,
    })
    expect(first.renderTileViewport).toEqual({
      colEnd: 511,
      colStart: 0,
      rowEnd: 191,
      rowStart: 0,
    })
    expect(first.residentHeaderRegion.range).toEqual({
      height: 96,
      width: 256,
      x: 256,
      y: 96,
    })
    expect(sameWindow.residentViewport).toBe(first.residentViewport)
    expect(sameWindow.visibleAddresses).toBe(first.visibleAddresses)
    expect(sameWindow.sceneRevision).toBe(1)
  })

  it('owns V3 header pane runtime state instead of letting the React hook allocate it', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })
    const panes = host.resolveHeaderPanes({
      columnWidths: { 0: 80, 1: 120 },
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: 80,
      frozenRowHeight: 30,
      getHeaderCellLocalBounds: (col: number, row: number): Rectangle => ({
        height: row === 0 ? 30 : 26,
        width: col === 0 ? 80 : 120,
        x: gridMetrics.rowMarkerWidth + (col === 0 ? 0 : 80),
        y: gridMetrics.headerHeight + (row === 0 ? 0 : 30),
      }),
      gridMetrics,
      hostClientHeight: 320,
      hostClientWidth: 520,
      hostReady: true,
      residentBodyPane: {
        contentOffset: { x: 33, y: 44 },
        surfaceSize: { width: 200, height: 56 },
      },
      residentHeaderItems: [
        [0, 0],
        [1, 0],
        [0, 1],
        [1, 1],
      ],
      residentHeaderRegion: {
        freezeCols: 1,
        freezeRows: 1,
        range: { height: 2, width: 2, x: 0, y: 0 },
        tx: 0,
        ty: 0,
      },
      residentViewport: { colEnd: 1, colStart: 0, rowEnd: 1, rowStart: 0 },
      rowHeights: { 0: 30, 1: 26 },
      sheetName: 'Sheet1',
    })

    expect(panes.map((pane) => pane.paneId)).toEqual(['corner-header', 'top-frozen', 'top-body', 'left-frozen', 'left-body'])
    expect(panes.find((pane) => pane.paneId === 'top-body')?.contentOffset).toEqual({ x: 33, y: 0 })
    expect(panes.find((pane) => pane.paneId === 'left-body')?.contentOffset).toEqual({ x: 0, y: 44 })
  })

  it('owns V3 render-tile pane runtime state instead of letting the React hook allocate it', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]

    const state = host.resolveRenderTilePanes({
      columnWidths: {},
      dprBucket: 1,
      engine: LOCAL_EMPTY_ENGINE,
      freezeCols: 0,
      freezeRows: 0,
      frozenColumnWidth: 0,
      frozenRowHeight: 0,
      gridMetrics,
      hostClientHeight: 400,
      hostClientWidth: 800,
      hostReady: true,
      renderTileSource: undefined,
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      residentViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      rowHeights: {},
      sceneRevision: 1,
      sheetId: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })

    expect(state.renderTilePanes).toHaveLength(1)
    expect(state.residentBodyPane?.tile.tileId).toBe(tileId)
    expect(state.tileReadiness).toMatchObject({
      exactHits: [tileId],
      misses: [],
      staleHits: [],
    })
    expect(host.tiles.residency.getExact(tileId)?.packet).toMatchObject({
      tileId,
      coord: {
        sheetId: 7,
      },
    })
  })

  it('resolves selection and restore scroll positions from runtime axes', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      freezeCols: 1,
      freezeRows: 1,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })
    host.updateAxes({
      columns: [
        { index: 0, size: 160 },
        { index: 2, size: 200 },
      ],
      rows: [{ index: 1, size: 30 }],
    })

    expect(
      host.resolveScrollForCellIntoView({
        cell: [3, 4],
        freezeCols: 1,
        freezeRows: 1,
        gridMetrics,
        scrollLeft: 0,
        scrollTop: 0,
        viewportHeight: 80,
        viewportWidth: 300,
      }),
    ).toEqual({
      scrollLeft: 310,
      scrollTop: 10,
    })
    expect(
      host.resolveScrollPositionForViewport({
        freezeCols: 1,
        freezeRows: 1,
        viewport: { colStart: 3, rowStart: 4 },
      }),
    ).toEqual({
      scrollLeft: 300,
      scrollTop: 50,
    })
  })
})
