import { describe, expect, it } from 'vitest'
import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import {
  GRID_TILE_PACKET_V3_MAGIC,
  createGridTilePacketV3,
  estimateGridTilePacketBytesV3,
  gridTilePacketRevisionTupleV3,
} from '../renderer-v3/tile-packet-v3.js'
import { packTileKey53 } from '../renderer-v3/tile-key.js'

describe('GridTilePacketV3', () => {
  it('derives fixed tile bounds and revisions from a packed numeric key', () => {
    const rectInstances = new Float32Array([1, 2, 3, 4])
    const textRuns = new Uint32Array([5, 6, 7])
    const dirtyMasks = new Uint32Array([1])
    const packet = createGridTilePacketV3({
      packetSeq: 11,
      materializedAtSeq: 12,
      tileKey: packTileKey53({ sheetOrdinal: 3, rowTile: 2, colTile: 4, dprBucket: 2 }),
      sheetId: 99,
      valueSeq: 13,
      styleSeq: 14,
      textSeq: 15,
      rectSeq: 16,
      axisSeqX: 17,
      axisSeqY: 18,
      freezeSeq: 19,
      glyphAtlasSeq: 20,
      cellCount: 21,
      rectInstanceCount: 1,
      textRunCount: 3,
      rectInstances,
      textRuns,
      dirtyMasks,
    })

    expect(packet).toMatchObject({
      magic: GRID_TILE_PACKET_V3_MAGIC,
      version: 1,
      packetSeq: 11,
      materializedAtSeq: 12,
      sheetId: 99,
      sheetOrdinal: 3,
      rowTile: 2,
      colTile: 4,
      rowStart: 2 * VIEWPORT_TILE_ROW_COUNT,
      rowEnd: 3 * VIEWPORT_TILE_ROW_COUNT - 1,
      colStart: 4 * VIEWPORT_TILE_COLUMN_COUNT,
      colEnd: 5 * VIEWPORT_TILE_COLUMN_COUNT - 1,
      dprBucket: 2,
      byteSize: rectInstances.byteLength + textRuns.byteLength + dirtyMasks.byteLength,
    })
    expect(gridTilePacketRevisionTupleV3(packet)).toEqual({
      valueSeq: 13,
      styleSeq: 14,
      textSeq: 15,
      rectSeq: 16,
      axisSeqX: 17,
      axisSeqY: 18,
      freezeSeq: 19,
      glyphAtlasSeq: 20,
    })
  })

  it('counts only typed-array payload bytes', () => {
    expect(
      estimateGridTilePacketBytesV3({
        rectInstances: new Float32Array(2),
        textRuns: new Uint32Array(3),
        dirtyLocalRows: new Uint32Array(4),
      }),
    ).toBe(36)
  })
})
