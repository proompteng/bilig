import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { buildWorkerRenderTileDeltaBatch } from '../worker-runtime-render-tile-delta.js'

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
