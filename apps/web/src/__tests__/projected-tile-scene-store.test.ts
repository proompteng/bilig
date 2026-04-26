import { describe, expect, it, vi } from 'vitest'
import { encodeRenderTileDeltaBatch, type RenderTileDeltaBatch, type RenderTileReplaceMutation } from '@bilig/worker-transport'
import { ProjectedTileSceneStore } from '../projected-tile-scene-store.js'

function createTileReplace(tileId: number, valuesVersion: number): RenderTileReplaceMutation {
  return {
    kind: 'tileReplace',
    tileId,
    coord: {
      sheetId: 7,
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
    textMetrics: new Float32Array([5, 6]),
    glyphRefs: new Uint32Array([9]),
    textRuns: [],
    textCount: 0,
    dirty: {
      rectSpans: [{ offset: 0, length: 1 }],
      textSpans: [],
      glyphSpans: [{ offset: 0, length: 1 }],
    },
  }
}

function createBatch(batchId: number, mutations: RenderTileDeltaBatch['mutations']): RenderTileDeltaBatch {
  return {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: 7,
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

    const staleChange = store.applyDelta(createBatch(1, [createTileReplace(101, 99)]))

    expect(staleChange.changedTileIds).toEqual([])
    expect(store.peekTile(101)?.version.values).toBe(2)
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
})
