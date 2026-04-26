import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT } from '@bilig/protocol'
import { unpackTileKey53, type TileKey53 } from './tile-key.js'
import type { TileRevisionTupleV3 } from './tile-residency.js'

export const GRID_TILE_PACKET_V3_MAGIC = 'bilig.grid.tile.v3'
export const GRID_TILE_PACKET_V3_VERSION = 1

export interface GridTilePacketRevisionTupleV3 extends TileRevisionTupleV3 {
  readonly glyphAtlasSeq: number
}

export interface GridTilePacketV3 extends GridTilePacketRevisionTupleV3 {
  readonly magic: typeof GRID_TILE_PACKET_V3_MAGIC
  readonly version: typeof GRID_TILE_PACKET_V3_VERSION
  readonly packetSeq: number
  readonly materializedAtSeq: number
  readonly tileKey: TileKey53
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly rowTile: number
  readonly colTile: number
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
  readonly dprBucket: number
  readonly cellCount: number
  readonly rectInstanceCount: number
  readonly textRunCount: number
  readonly byteSize: number
  readonly rectInstances: Float32Array
  readonly textRuns: Uint32Array | Float32Array
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
}

export interface CreateGridTilePacketV3Input extends GridTilePacketRevisionTupleV3 {
  readonly packetSeq: number
  readonly materializedAtSeq: number
  readonly tileKey: TileKey53
  readonly sheetId: number
  readonly cellCount: number
  readonly rectInstanceCount: number
  readonly textRunCount: number
  readonly rectInstances: Float32Array
  readonly textRuns: Uint32Array | Float32Array
  readonly byteSize?: number | undefined
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
}

export function createGridTilePacketV3(input: CreateGridTilePacketV3Input): GridTilePacketV3 {
  const key = unpackTileKey53(input.tileKey)
  return {
    magic: GRID_TILE_PACKET_V3_MAGIC,
    version: GRID_TILE_PACKET_V3_VERSION,
    packetSeq: input.packetSeq,
    materializedAtSeq: input.materializedAtSeq,
    tileKey: input.tileKey,
    sheetId: input.sheetId,
    sheetOrdinal: key.sheetOrdinal,
    rowTile: key.rowTile,
    colTile: key.colTile,
    rowStart: key.rowTile * VIEWPORT_TILE_ROW_COUNT,
    rowEnd: Math.min(MAX_ROWS - 1, (key.rowTile + 1) * VIEWPORT_TILE_ROW_COUNT - 1),
    colStart: key.colTile * VIEWPORT_TILE_COLUMN_COUNT,
    colEnd: Math.min(MAX_COLS - 1, (key.colTile + 1) * VIEWPORT_TILE_COLUMN_COUNT - 1),
    dprBucket: key.dprBucket,
    valueSeq: input.valueSeq,
    styleSeq: input.styleSeq,
    textSeq: input.textSeq,
    rectSeq: input.rectSeq,
    axisSeqX: input.axisSeqX,
    axisSeqY: input.axisSeqY,
    freezeSeq: input.freezeSeq,
    glyphAtlasSeq: input.glyphAtlasSeq,
    cellCount: input.cellCount,
    rectInstanceCount: input.rectInstanceCount,
    textRunCount: input.textRunCount,
    byteSize: input.byteSize ?? estimateGridTilePacketBytesV3(input),
    rectInstances: input.rectInstances,
    textRuns: input.textRuns,
    ...(input.dirtyLocalRows ? { dirtyLocalRows: input.dirtyLocalRows } : {}),
    ...(input.dirtyLocalCols ? { dirtyLocalCols: input.dirtyLocalCols } : {}),
    ...(input.dirtyMasks ? { dirtyMasks: input.dirtyMasks } : {}),
  }
}

export function gridTilePacketRevisionTupleV3(packet: GridTilePacketV3): GridTilePacketRevisionTupleV3 {
  return {
    axisSeqX: packet.axisSeqX,
    axisSeqY: packet.axisSeqY,
    freezeSeq: packet.freezeSeq,
    glyphAtlasSeq: packet.glyphAtlasSeq,
    rectSeq: packet.rectSeq,
    styleSeq: packet.styleSeq,
    textSeq: packet.textSeq,
    valueSeq: packet.valueSeq,
  }
}

export function estimateGridTilePacketBytesV3(input: {
  readonly rectInstances: Float32Array
  readonly textRuns: Uint32Array | Float32Array
  readonly dirtyLocalRows?: Uint32Array | undefined
  readonly dirtyLocalCols?: Uint32Array | undefined
  readonly dirtyMasks?: Uint32Array | undefined
}): number {
  return (
    input.rectInstances.byteLength +
    input.textRuns.byteLength +
    (input.dirtyLocalRows?.byteLength ?? 0) +
    (input.dirtyLocalCols?.byteLength ?? 0) +
    (input.dirtyMasks?.byteLength ?? 0)
  )
}
