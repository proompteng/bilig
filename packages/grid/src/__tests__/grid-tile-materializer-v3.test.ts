import { describe, expect, test } from 'vitest'
import { ValueTag, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
import { getGridMetrics } from '../gridMetrics.js'
import type { GridEngineLike } from '../grid-engine.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { materializeGridRenderTileV3 } from '../renderer-v3/grid-tile-materializer.js'
import { GRID_TILE_PACKET_V3_MAGIC } from '../renderer-v3/tile-packet-v3.js'

function createCellSnapshot(address: string, value: CellSnapshot['value'], styleId: string | undefined = 'style-1'): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value,
    version: 0,
    ...(styleId ? { styleId } : {}),
  }
}

function makeEngine(cells: Record<string, CellSnapshot>, styles: Record<string, CellStyleRecord> = {}): GridEngineLike {
  return {
    getCell: (_sheetName, address) => cells[address] ?? createCellSnapshot(address, { tag: ValueTag.Empty }, undefined),
    getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
    subscribeCells: () => () => undefined,
    workbook: {
      getSheet: () => undefined,
    },
  }
}

describe('renderer-v3 grid tile materializer', () => {
  test('materializes fixed content tiles with native v3 packets instead of v2 scene packets', () => {
    const gridMetrics = getGridMetrics()
    const tile = materializeGridRenderTileV3({
      axisSeqX: 5,
      axisSeqY: 6,
      cameraSeq: 7,
      columnWidths: {},
      dprBucket: 2,
      engine: makeEngine(
        {
          A1: createCellSnapshot('A1', { tag: ValueTag.String, value: 'alpha' }),
        },
        {
          'style-1': {
            fill: { backgroundColor: '#ff0000' },
          },
        },
      ),
      freezeSeq: 8,
      glyphAtlasSeq: 9,
      gridMetrics,
      materializedAtSeq: 10,
      packetSeq: 11,
      rectSeq: 12,
      rowHeights: {},
      sheetId: 3,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 13,
      textSeq: 14,
      valueSeq: 15,
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT - 1,
        rowStart: 0,
      },
    })

    expect(tile.packet?.magic).toBe(GRID_TILE_PACKET_V3_MAGIC)
    expect(tile.tileId).toBe(tile.packet?.tileKey)
    expect(tile.coord).toMatchObject({ colTile: 0, dprBucket: 2, rowTile: 0, sheetId: 3 })
    expect(tile.version).toEqual({ axisX: 5, axisY: 6, freeze: 8, styles: 13, text: 14, values: 15 })
    expect(tile.bounds).toEqual({ colEnd: VIEWPORT_TILE_COLUMN_COUNT - 1, colStart: 0, rowEnd: VIEWPORT_TILE_ROW_COUNT - 1, rowStart: 0 })
    expect(tile.textRuns).toContainEqual(expect.objectContaining({ text: 'alpha', x: 0, y: 0 }))
    expect(tile.rectCount).toBeGreaterThan(0)
    expect(tile.rectInstances.length).toBeGreaterThanOrEqual(tile.rectCount * 20)
    expect(tile.packet && 'key' in tile.packet).toBe(false)
  })

  test('local fixed tile generation returns one v3 packet per fixed protocol tile', () => {
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 4,
      columnWidths: {},
      dprBucket: 1,
      engine: makeEngine({
        A1: createCellSnapshot('A1', { tag: ValueTag.String, value: 'visible' }),
      }),
      generation: 21,
      gridMetrics: getGridMetrics(),
      rowHeights: {},
      sheetId: 2,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        colEnd: VIEWPORT_TILE_COLUMN_COUNT,
        colStart: 0,
        rowEnd: VIEWPORT_TILE_ROW_COUNT,
        rowStart: 0,
      },
    })

    expect(tiles).toHaveLength(4)
    expect(tiles.map((tile) => tile.packet?.magic)).toEqual([
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
      GRID_TILE_PACKET_V3_MAGIC,
    ])
  })
})
