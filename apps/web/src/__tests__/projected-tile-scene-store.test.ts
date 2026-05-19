import { describe, expect, it, vi } from 'vitest'
import {
  encodeRenderTileDeltaBatch,
  encodeWorkbookDeltaBatchV3,
  type RenderTileDeltaBatch,
  type RenderTileReplaceMutation,
} from '@bilig/worker-transport'
import { ValueTag } from '@bilig/protocol'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'
import { ProjectedTileSceneStore } from '../projected-tile-scene-store.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'

const LOCAL_OPTIMISTIC_CELL_VISUAL_DIRTY_MASK =
  DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border

function createTileReplace(tileId: number, valuesVersion: number, sheetOrdinal = 7, sheetId = 7): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId,
    coord: {
      sheetId,
      sheetOrdinal,
      paneKind: 'body',
      rowTile: tileId,
      colTile: 0,
      dprBucket: 1,
    },
    version: {
      axisX: 1,
      axisY: 1,
      values: valuesVersion,
      styles: valuesVersion,
      text: valuesVersion,
      freeze: 1,
    },
    bounds: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 63 },
    rectInstances: new Float32Array([1, 2, 3, 4]),
    rectCount: 1,
    rectSignature: `rect:${sheetId}:${sheetOrdinal}:${tileId}:${valuesVersion}`,
    textMetrics: new Float32Array([5, 6]),
    glyphRefs: new Uint32Array([9]),
    textRuns: [],
    textCount: 0,
    textSignature: `text:${sheetId}:${sheetOrdinal}:${tileId}:${valuesVersion}`,
    dirty: {
      rectSpans: [{ offset: 0, length: 1 }],
      textSpans: [],
      glyphSpans: [{ offset: 0, length: 1 }],
    },
    dirtyLocalRows: new Uint32Array([0, 0]),
    dirtyLocalCols: new Uint32Array([1, 1]),
    dirtyMasks: new Uint32Array([5]),
  }
}

function createBatch(batchId: number, mutations: RenderTileDeltaBatch['mutations']): RenderTileDeltaBatch {
  return {
    magic: 'bilig.render.tile.delta',
    version: 4,
    sheetId: 7,
    sheetOrdinal: 7,
    batchId,
    cameraSeq: batchId + 10,
    mutations,
  }
}

describe('ProjectedTileSceneStore', () => {
  it('applies tile replacements and ignores stale batches', () => {
    const store = new ProjectedTileSceneStore()
    const change = store.applyDelta(createBatch(2, [createTileReplace(101, 2)]))

    expect(change).toMatchObject({
      batchId: 2,
      cameraSeq: 12,
      changedTileIds: [101],
      invalidatedTileIds: [],
      structural: false,
    })
    expect(store.peekTile(101)).toMatchObject({ tileId: 101, rectCount: 1, lastBatchId: 2 })
    expect(store.peekTile(101)?.rectSignature).toBe('rect:7:7:101:2')
    expect(store.peekTile(101)?.textSignature).toBe('text:7:7:101:2')
    expect(store.peekTile(101)?.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(store.peekTile(101)?.dirtyLocalCols).toEqual(new Uint32Array([1, 1]))
    expect(store.peekTile(101)?.dirtyMasks).toEqual(new Uint32Array([5]))

    const staleChange = store.applyDelta(createBatch(1, [createTileReplace(101, 99)]))

    expect(staleChange.changedTileIds).toEqual([])
    expect(store.peekTile(101)?.version.values).toBe(2)
  })

  it('ignores older camera sequences for the current batch id', () => {
    const store = new ProjectedTileSceneStore()
    store.applyDelta({ ...createBatch(2, [createTileReplace(101, 2)]), cameraSeq: 20 })

    const staleCameraChange = store.applyDelta({ ...createBatch(2, [createTileReplace(101, 99)]), cameraSeq: 19 })

    expect(staleCameraChange.changedTileIds).toEqual([])
    expect(store.peekTile(101)?.version.values).toBe(2)
  })

  it('does not let one sheet camera sequence suppress another sheet current-batch delta', () => {
    const store = new ProjectedTileSceneStore()
    store.applyDelta({
      ...createBatch(2, [createTileReplace(101, 2, 1, 7)]),
      cameraSeq: 20,
      sheetId: 7,
      sheetOrdinal: 1,
    })

    const otherSheetChange = store.applyDelta({
      ...createBatch(2, [createTileReplace(202, 2, 2, 8)]),
      cameraSeq: 3,
      sheetId: 8,
      sheetOrdinal: 2,
    })

    expect(otherSheetChange.changedTileIds).toEqual([202])
    expect(store.peekTile(101)).toMatchObject({ tileId: 101, coord: expect.objectContaining({ sheetId: 7, sheetOrdinal: 1 }) })
    expect(store.peekTile(202)).toMatchObject({ tileId: 202, coord: expect.objectContaining({ sheetId: 8, sheetOrdinal: 2 }) })
  })

  it('clears stale sheet tiles for structural batches before applying replacements', () => {
    const store = new ProjectedTileSceneStore()
    store.applyDelta(createBatch(2, [createTileReplace(101, 2)]))

    const change = store.applyDelta(
      createBatch(3, [{ kind: 'axis', axis: 'col', changedStart: 0, changedEnd: 3, axisVersion: 4 }, createTileReplace(202, 3)]),
    )

    expect(change.structural).toBe(true)
    expect(change.invalidatedTileIds).toEqual([101])
    expect(change.changedTileIds).toEqual([202])
    expect(store.peekTile(101)).toBeNull()
    expect(store.peekTile(202)).toMatchObject({ tileId: 202, lastBatchId: 3 })
  })

  it('does not invalidate tiles from a different sheet that reuses an ordinal', () => {
    const store = new ProjectedTileSceneStore()
    store.applyDelta({
      ...createBatch(2, [createTileReplace(101, 2, 7, 7)]),
      sheetId: 7,
      sheetOrdinal: 7,
    })

    const change = store.applyDelta({
      ...createBatch(3, [{ kind: 'axis', axis: 'col', changedStart: 0, changedEnd: 3, axisVersion: 4 }]),
      sheetId: 99,
      sheetOrdinal: 7,
    })

    expect(change.invalidatedTileIds).toEqual([])
    expect(store.peekTile(101)).toMatchObject({ tileId: 101, coord: expect.objectContaining({ sheetId: 7, sheetOrdinal: 7 }) })
  })

  it('rejects partial cell-run mutations by invalidating the tile instead of keeping stale visuals', () => {
    const store = new ProjectedTileSceneStore()
    store.applyDelta(createBatch(2, [createTileReplace(101, 2)]))

    const change = store.applyDelta(
      createBatch(3, [
        {
          kind: 'cellRuns',
          runs: [
            {
              colEnd: 1,
              colStart: 1,
              glyphSpan: { length: 0, offset: 0 },
              rectSpan: { length: 1, offset: 0 },
              row: 1,
              textSpan: { length: 0, offset: 0 },
            },
          ],
          tileId: 101,
          version: {
            axisX: 1,
            axisY: 1,
            freeze: 1,
            styles: 3,
            text: 3,
            values: 3,
          },
        },
      ]),
    )

    expect(change.changedTileIds).toEqual([])
    expect(change.invalidatedTileIds).toEqual([101])
    expect(store.peekTile(101)).toBeNull()
  })

  it('subscribes to worker render tile deltas and decodes incoming batches', () => {
    let emit: ((bytes: Uint8Array) => void) | null = null
    const unsubscribeWorker = vi.fn()
    const client = {
      subscribeRenderTileDeltas: vi.fn((_subscription, listener: (bytes: Uint8Array) => void) => {
        emit = listener
        return unsubscribeWorker
      }),
    }
    const store = new ProjectedTileSceneStore(client)
    const listener = vi.fn()

    const unsubscribe = store.subscribe(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      listener,
    )
    emit?.(encodeRenderTileDeltaBatch(createBatch(2, [createTileReplace(101, 2)])))

    expect(listener).toHaveBeenCalledWith(expect.objectContaining({ changedTileIds: [101] }))
    expect(store.peekTile(101)).toMatchObject({ tileId: 101 })

    unsubscribe()
    expect(unsubscribeWorker).toHaveBeenCalledTimes(1)
  })

  it('indexes subscribed sheets by ordinal when sheet ids differ from order', () => {
    const store = new ProjectedTileSceneStore({
      subscribeRenderTileDeltas(_subscription, listener) {
        listener(
          encodeRenderTileDeltaBatch({
            magic: 'bilig.render.tile.delta',
            version: 4,
            sheetId: 99,
            sheetOrdinal: 2,
            batchId: 4,
            cameraSeq: 5,
            mutations: [createTileReplace(202, 4, 2, 99)],
          }),
        )
        return () => undefined
      },
    })

    store.subscribe(
      {
        sheetId: 99,
        sheetOrdinal: 2,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
      () => undefined,
    )

    expect(store.peekTile(202)?.coord).toMatchObject({ sheetId: 99, sheetOrdinal: 2 })
    store.dropSheets(['Sheet1'])
    expect(store.peekTile(202)).toBeNull()
  })

  it('drops only tiles whose sheet id and ordinal both match the subscribed sheet identity', () => {
    const store = new ProjectedTileSceneStore({
      subscribeRenderTileDeltas(_subscription, listener) {
        listener(
          encodeRenderTileDeltaBatch({
            magic: 'bilig.render.tile.delta',
            version: 4,
            sheetId: 99,
            sheetOrdinal: 2,
            batchId: 4,
            cameraSeq: 5,
            mutations: [createTileReplace(202, 4, 2, 99)],
          }),
        )
        listener(
          encodeRenderTileDeltaBatch({
            magic: 'bilig.render.tile.delta',
            version: 4,
            sheetId: 7,
            sheetOrdinal: 2,
            batchId: 5,
            cameraSeq: 6,
            mutations: [createTileReplace(303, 5, 2, 7)],
          }),
        )
        return () => undefined
      },
    })

    store.subscribe(
      {
        sheetId: 99,
        sheetOrdinal: 2,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
      () => undefined,
    )

    store.dropSheets(['Sheet1'])

    expect(store.peekTile(202)).toBeNull()
    expect(store.peekTile(303)).toMatchObject({ tileId: 303, coord: expect.objectContaining({ sheetId: 7, sheetOrdinal: 2 }) })
  })
})

describe('ProjectedViewportStore render delta source bridge', () => {
  it('publishes local optimistic workbook deltas for projected cell snapshots', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    store.setCellSnapshot({
      address: 'B2',
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 17 },
      version: 12,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: expect.objectContaining({ cellRanges: new Uint32Array([1, 1, 1, 1, LOCAL_OPTIMISTIC_CELL_VISUAL_DIRTY_MASK]) }),
        seq: 1,
        sheetId: 7,
        sheetOrdinal: 3,
        source: 'localOptimistic',
        valueSeq: 12,
      }),
    )

    unsubscribe()
  })

  it('publishes deltas from the accepted projected snapshot after optimistic normalization', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    store.setCellSnapshot({
      address: 'B2',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      formula: '1+1',
      input: '=1+1',
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 2 },
      version: 8,
    })

    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    store.setCellSnapshot({
      address: 'B2',
      flags: 0,
      formula: '1+1',
      input: '=1+1',
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 3 },
      version: 7,
    })

    expect(store.getCell('Sheet1', 'B2')).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 8,
    })
    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        calcSeq: 8,
        styleSeq: 8,
        valueSeq: 8,
      }),
    )

    unsubscribe()
  })

  it('keeps local optimistic workbook deltas newer than observed render batches', async () => {
    let emitRenderDelta: ((bytes: Uint8Array) => void) | null = null
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: vi.fn((_subscription, listener: (bytes: Uint8Array) => void) => {
        emitRenderDelta = listener
        return () => undefined
      }),
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 32,
        rowEnd: 63,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    await vi.waitFor(() => {
      expect(emitRenderDelta).not.toBeNull()
    })
    emitRenderDelta?.(
      encodeRenderTileDeltaBatch({
        ...createBatch(42, [createTileReplace(101, 42, 3, 7)]),
        sheetOrdinal: 3,
      }),
    )
    const unsubscribe = store.subscribeWorkbookDeltas(listener)

    store.setCellSnapshot({
      address: 'D53',
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Empty },
      version: 13,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: expect.objectContaining({ cellRanges: new Uint32Array([52, 52, 3, 3, LOCAL_OPTIMISTIC_CELL_VISUAL_DIRTY_MASK]) }),
        seq: 43,
        sheetId: 7,
        sheetOrdinal: 3,
        source: 'localOptimistic',
      }),
    )

    unsubscribe()
  })

  it('does not publish optimistic workbook deltas before render tile sheet identity is known', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    store.setCellSnapshot({
      address: 'B2',
      flags: 0,
      sheetName: 'Sheet1',
      value: { tag: ValueTag.Number, value: 17 },
      version: 12,
    })

    expect(listener).not.toHaveBeenCalled()

    unsubscribe()
  })

  it('publishes optimistic workbook deltas after runtime sheet identities are registered', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    store.setCellSnapshot({
      address: 'D53',
      flags: 0,
      input: 'Month 1',
      sheetName: 'Sheet1',
      value: { tag: ValueTag.String, stringId: 0, value: 'Month 1' },
      version: 12,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: expect.objectContaining({ cellRanges: new Uint32Array([52, 52, 3, 3, LOCAL_OPTIMISTIC_CELL_VISUAL_DIRTY_MASK]) }),
        sheetId: 7,
        sheetOrdinal: 3,
        source: 'localOptimistic',
      }),
    )

    unsubscribe()
  })

  it('publishes style-aware local optimistic workbook deltas for styled projected cell snapshots', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    store.setCellSnapshot({
      address: 'B2',
      flags: 0,
      sheetName: 'Sheet1',
      styleId: 'bold',
      value: { tag: ValueTag.Number, value: 17 },
      version: 12,
    })

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: expect.objectContaining({ cellRanges: new Uint32Array([1, 1, 1, 1, LOCAL_OPTIMISTIC_CELL_VISUAL_DIRTY_MASK]) }),
      }),
    )

    unsubscribe()
  })

  it('publishes local optimistic workbook deltas for projected axis size updates', () => {
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: () => () => undefined,
    })
    const listener = vi.fn()

    store.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribe = store.subscribeWorkbookDeltas(listener)

    store.setColumnWidth('Sheet1', 2, 144)
    store.setRowHeight('Sheet1', 4, 40)

    expect(listener).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        axisSeqX: 1,
        dirty: expect.objectContaining({
          axisX: new Uint32Array([2, 2, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
          axisY: new Uint32Array(),
        }),
        seq: 1,
        source: 'localOptimistic',
      }),
    )
    expect(listener).toHaveBeenNthCalledWith(
      2,
      expect.objectContaining({
        axisSeqY: 2,
        dirty: expect.objectContaining({
          axisX: new Uint32Array(),
          axisY: new Uint32Array([4, 4, DirtyMaskV3.AxisY | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        }),
        seq: 2,
        source: 'localOptimistic',
      }),
    )

    unsubscribe()
  })

  it('decodes workbook deltas for the grid runtime dirty-tile coordinator', () => {
    let emit: ((bytes: Uint8Array) => void) | null = null
    const unsubscribeWorker = vi.fn()
    const store = new ProjectedViewportStore({
      subscribeRenderTileDeltas: () => () => undefined,
      subscribeViewportPatches: () => () => undefined,
      subscribeWorkbookDeltas: vi.fn((listener: (bytes: Uint8Array) => void) => {
        emit = listener
        return unsubscribeWorker
      }),
    })
    const listener = vi.fn()

    const unsubscribe = store.subscribeWorkbookDeltas(listener)
    emit?.(
      encodeWorkbookDeltaBatchV3({
        calcSeq: 3,
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([0, 0, 0, 0, 1]),
        },
        freezeSeq: 1,
        magic: 'bilig.workbook.delta.v3',
        seq: 11,
        sheetId: 7,
        sheetOrdinal: 7,
        source: 'workerAuthoritative',
        styleSeq: 2,
        valueSeq: 3,
        version: 1,
        axisSeqX: 4,
        axisSeqY: 5,
      }),
    )

    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: expect.objectContaining({ cellRanges: new Uint32Array([0, 0, 0, 0, 1]) }),
        seq: 11,
        sheetId: 7,
        sheetOrdinal: 7,
      }),
    )

    unsubscribe()
    expect(unsubscribeWorker).toHaveBeenCalledTimes(1)
  })
})
