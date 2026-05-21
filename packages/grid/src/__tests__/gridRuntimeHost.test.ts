import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { Rectangle } from '../gridTypes.js'
import { createGridSelection } from '../gridSelection.js'
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

  it('owns render tile bridge snapshots and subscriptions', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })
    const snapshots: unknown[] = []
    const unsubscribe = host.subscribeRenderTileBridgeState(() => {
      snapshots.push(host.snapshotRenderTileBridgeState())
    })

    host.noteRenderTileDelta()
    host.noteWorkbookDeltaDamage()
    unsubscribe()
    host.noteLocalRenderTileFallbackInvalidation()

    expect(snapshots).toEqual([
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 1,
      },
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 2,
      },
    ])
    expect(host.snapshotRenderTileBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
  })

  it('owns interaction overlay snapshots outside React state', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })
    const snapshots: unknown[] = []
    const unsubscribe = host.interactionOverlays.subscribe(() => {
      snapshots.push(host.interactionOverlays.snapshot())
    })

    host.interactionOverlays.syncSelectedCell({ selectedCol: 3, selectedRow: 4 })
    host.interactionOverlays.setFillPreviewRange({ height: 1, width: 2, x: 3, y: 4 })
    host.interactionOverlays.setIsFillHandleDragging(true)
    unsubscribe()
    host.interactionOverlays.setIsRangeMoveDragging(true)

    expect(snapshots).toHaveLength(3)
    expect(host.interactionOverlays.snapshot()).toMatchObject({
      fillPreviewRange: { height: 1, width: 2, x: 3, y: 4 },
      isFillHandleDragging: true,
      isRangeMoveDragging: true,
    })
    expect(host.interactionOverlays.snapshot().gridSelection.current?.cell).toEqual([3, 4])
    expect(
      host.interactionOverlays.resolveState({
        activeResizeColumn: null,
        activeResizeRow: null,
        getCellLocalBounds: (col, row) =>
          row === 4 && col >= 3 && col <= 4 ? { height: 10, width: 40, x: col === 3 ? 30 : 70, y: 20 } : undefined,
        hasColumnResizePreview: false,
        hasRowResizePreview: false,
        isEditingCell: false,
        snapshot: host.interactionOverlays.snapshot(),
        visibleRange: { height: 10, width: 10, x: 0, y: 0 },
      }),
    ).toMatchObject({
      fillPreviewBounds: { height: 10, width: 80, x: 30, y: 20 },
      requiresLiveViewportState: true,
      selectionRange: { height: 1, width: 1, x: 3, y: 4 },
    })
  })

  it('normalizes detached active cells before overlay state can render stale area chrome', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    host.interactionOverlays.setGridSelection({
      ...createGridSelection(6, 7),
      current: {
        cell: [6, 7],
        range: { height: 3, width: 3, x: 1, y: 1 },
        rangeStack: [],
      },
    })

    expect(host.interactionOverlays.snapshot().gridSelection).toEqual(createGridSelection(6, 7))
  })

  it('owns input interaction refs for the workbook runtime lifetime', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })

    const interactionState = host.input.interactionState
    interactionState.pendingPointerCellRef.current = [1, 2]
    host.input.syncFillPreviewRange({ height: 1, width: 2, x: 3, y: 4 })

    expect(host.input.interactionState).toBe(interactionState)
    expect(host.input.pendingPointerCellRef.current).toEqual([1, 2])
    expect(host.input.fillPreviewRangeRef.current).toEqual({ height: 1, width: 2, x: 3, y: 4 })
  })

  it('disposes host-owned input and renderer subscriptions together', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 60,
      viewportWidth: 250,
    })
    const fillCleanup = vi.fn()
    const renderDisconnect = vi.fn()
    const viewportScrollDispose = vi.spyOn(host.viewportScroll, 'dispose')

    host.input.fillHandleCleanupRef.current = fillCleanup
    host.syncRenderTileConnections({
      dprBucket: 1,
      engine: LOCAL_EMPTY_ENGINE,
      needsLocalCellInvalidation: false,
      renderTileSource: {
        peekRenderTile: () => null,
        subscribeRenderTileDeltas: () => renderDisconnect,
      },
      renderTileViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      sheetId: 1,
      sheetName: 'Sheet1',
      sheetOrdinal: 0,
      visibleAddresses: [],
    })

    host.dispose()
    host.dispose()

    expect(fillCleanup).toHaveBeenCalledTimes(1)
    expect(renderDisconnect).toHaveBeenCalledTimes(1)
    expect(viewportScrollDispose).toHaveBeenCalledTimes(1)
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
      visibleRegion: {
        freezeCols: 2,
        freezeRows: 1,
        range: { height: 8, width: 10, x: 260, y: 110 },
        tx: 0,
        ty: 0,
      },
    })
    host.viewportResidency.invalidateScene()
    const sameWindow = host.resolveViewportResidency({
      freezeCols: 2,
      freezeRows: 1,
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

  it('owns viewport residency scene revision subscriptions', () => {
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
    const snapshots: number[] = []
    const unsubscribe = host.subscribeViewportResidencySceneRevision(() => {
      snapshots.push(host.snapshotViewportResidencySceneRevision())
    })

    expect(host.snapshotViewportResidencySceneRevision()).toBe(0)
    host.viewportResidency.invalidateScene()
    host.viewportResidency.invalidateScene()
    unsubscribe()
    host.viewportResidency.invalidateScene()

    expect(snapshots).toEqual([1, 2])
    expect(host.snapshotViewportResidencySceneRevision()).toBe(3)
  })

  it('reconciles viewport residency invalidation subscriptions through the host', () => {
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
    let invalidateScene: (() => void) | null = null
    const unsubscribe = vi.fn()
    const subscribeCells = vi.fn((_sheetName: string, _addresses: readonly string[], listener: () => void) => {
      invalidateScene = listener
      return unsubscribe
    })
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      subscribeCells,
    }

    host.syncViewportResidencyInvalidation({
      engine,
      sheetName: 'Sheet1',
      shouldUseRemoteRenderTileSource: false,
      visibleAddresses: ['A1', 'B2'],
    })
    host.syncViewportResidencyInvalidation({
      engine,
      sheetName: 'Sheet1',
      shouldUseRemoteRenderTileSource: false,
      visibleAddresses: ['A1', 'B2'],
    })

    expect(subscribeCells).toHaveBeenCalledTimes(1)
    expect(unsubscribe).not.toHaveBeenCalled()

    invalidateScene?.()

    expect(host.snapshotViewportResidencySceneRevision()).toBe(1)

    host.syncViewportResidencyInvalidation({
      engine,
      sheetName: 'Sheet1',
      shouldUseRemoteRenderTileSource: true,
      visibleAddresses: ['A1', 'B2'],
    })

    expect(unsubscribe).toHaveBeenCalledTimes(1)
    host.disconnectViewportResidencyInvalidation()
    expect(unsubscribe).toHaveBeenCalledTimes(1)
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

  it('owns render-tile revision bridge state outside the React hook', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 80,
      viewportWidth: 300,
    })

    expect(host.snapshotRenderTileBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 0,
    })

    expect(host.noteLocalRenderTileFallbackInvalidation()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 0,
    })
    expect(host.noteWorkbookDeltaDamage()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })
    expect(host.noteRenderTileDelta()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
    expect(host.snapshotRenderTileBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 2,
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

  it('does not autoscroll when the selected cell is already partially visible', () => {
    const host = new GridRuntimeHost({
      columnCount: 1000,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 1000,
      viewportHeight: 116,
      viewportWidth: 300,
    })

    expect(
      host.resolveScrollForCellIntoView({
        cell: [2, 9],
        gridMetrics,
        scrollLeft: 0,
        scrollTop: 0,
        viewportHeight: 116,
        viewportWidth: 300,
      }),
    ).toEqual({
      scrollLeft: 0,
      scrollTop: 0,
    })

    expect(
      host.resolveScrollForCellIntoView({
        cell: [3, 10],
        gridMetrics,
        scrollLeft: 0,
        scrollTop: 0,
        viewportHeight: 116,
        viewportWidth: 300,
      }),
    ).toEqual({
      scrollLeft: 150,
      scrollTop: 14,
    })
  })
})
