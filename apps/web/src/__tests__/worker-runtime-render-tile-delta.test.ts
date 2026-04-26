import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createGridTileKeyV2, packGridScenePacketV2 } from '../../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { buildRenderTileDeltaBatchFromResidentPaneScenes, buildWorkerRenderTileDeltaBatch } from '../worker-runtime-render-tile-delta.js'

const emptyCell: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'A1',
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
}

const engine = {
  workbook: {
    getSheet: () => undefined,
  },
  getCell: () => emptyCell,
  getCellStyle: () => undefined,
  getColumnAxisEntries: () => [],
  getRowAxisEntries: () => [],
  subscribeCells: () => () => undefined,
  getLastMetrics: () => ({ batchId: 3 }),
} as const

describe('worker-runtime-render-tile-delta', () => {
  it('converts resident pane scenes into renderer tile replacement deltas', () => {
    const viewport = { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 63 }
    const packedScene = packGridScenePacketV2({
      generation: 3,
      requestSeq: 11,
      cameraSeq: 19,
      sheetName: 'Sheet1',
      paneId: 'body',
      viewport,
      surfaceSize: { width: 200, height: 100 },
      key: createGridTileKeyV2({
        sheetName: 'Sheet1',
        paneId: 'body',
        viewport,
        rowTile: 2,
        colTile: 4,
        axisVersionX: 5,
        axisVersionY: 6,
        valueVersion: 7,
        styleVersion: 8,
        textEpoch: 9,
        freezeVersion: 10,
        dprBucket: 2,
      }),
      gpuScene: { fillRects: [], borderRects: [] },
      textScene: { items: [] },
    })

    const batch = buildRenderTileDeltaBatchFromResidentPaneScenes({
      sheetId: 42,
      batchId: 7,
      cameraSeq: 19,
      scenes: [
        {
          generation: 3,
          paneId: 'body',
          viewport,
          surfaceSize: { width: 200, height: 100 },
          packedScene,
        },
      ],
    })

    expect(batch).toMatchObject({
      magic: 'bilig.render.tile.delta',
      version: 1,
      sheetId: 42,
      batchId: 7,
      cameraSeq: 19,
    })
    expect(batch.mutations).toHaveLength(1)
    expect(batch.mutations[0]).toMatchObject({
      kind: 'tileReplace',
      coord: {
        sheetId: 42,
        paneKind: 'body',
        rowTile: 2,
        colTile: 4,
        dprBucket: 2,
      },
      version: {
        axisX: 5,
        axisY: 6,
        values: 7,
        styles: 8,
        text: 9,
        freeze: 10,
      },
      bounds: viewport,
    })
  })

  it('materializes fixed content tiles instead of frozen pane scene duplicates', () => {
    const batch = buildWorkerRenderTileDeltaBatch({
      engine,
      generation: 4,
      subscription: {
        sheetId: 7,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 33,
        colStart: 0,
        colEnd: 129,
        dprBucket: 2,
        cameraSeq: 17,
      },
    })

    const replacements = batch.mutations.filter((mutation) => mutation.kind === 'tileReplace')

    expect(batch).toMatchObject({ batchId: 3, cameraSeq: 17, sheetId: 7 })
    expect(replacements).toHaveLength(4)
    expect(replacements.map((mutation) => mutation.coord)).toEqual([
      expect.objectContaining({ paneKind: 'body', rowTile: 0, colTile: 0 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 0, colTile: 1 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 1, colTile: 0 }),
      expect.objectContaining({ paneKind: 'body', rowTile: 1, colTile: 1 }),
    ])
    expect(replacements.map((mutation) => mutation.bounds)).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ])
  })
})
