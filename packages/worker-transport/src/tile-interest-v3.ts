import { BinaryProtocolError, BinaryReader, BinaryWriter } from '@bilig/binary-protocol'

export type TileKey53 = number

export type TileInterestReasonV3 = 'scroll' | 'sheetSwitch' | 'mutation' | 'viewportRestore' | 'prefetch'

export interface TileInterestBatchV3 {
  readonly magic: 'bilig.tile.interest.v3'
  readonly version: 1
  readonly seq: number
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly cameraSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly visibleTileKeys: readonly TileKey53[]
  readonly warmTileKeys: readonly TileKey53[]
  readonly pinnedTileKeys: readonly TileKey53[]
  readonly reason: TileInterestReasonV3
}

const TILE_INTEREST_V3_MAGIC = 0x54495633
const TILE_INTEREST_V3_VERSION = 1

const REASON_TAGS: Record<TileInterestReasonV3, number> = {
  scroll: 1,
  sheetSwitch: 2,
  mutation: 3,
  viewportRestore: 4,
  prefetch: 5,
}

const REASON_BY_TAG = new Map<number, TileInterestReasonV3>([
  [1, 'scroll'],
  [2, 'sheetSwitch'],
  [3, 'mutation'],
  [4, 'viewportRestore'],
  [5, 'prefetch'],
])

export function encodeTileInterestBatchV3(batch: TileInterestBatchV3): Uint8Array {
  const writer = new BinaryWriter()
  writer.u32(TILE_INTEREST_V3_MAGIC)
  writer.u32(TILE_INTEREST_V3_VERSION)
  writer.u32(batch.seq)
  writer.u32(batch.sheetId)
  writer.u32(batch.sheetOrdinal)
  writer.u32(batch.cameraSeq)
  writer.u32(batch.axisSeqX)
  writer.u32(batch.axisSeqY)
  writer.u32(batch.freezeSeq)
  encodeTileKeys(writer, batch.visibleTileKeys)
  encodeTileKeys(writer, batch.warmTileKeys)
  encodeTileKeys(writer, batch.pinnedTileKeys)
  writer.u8(REASON_TAGS[batch.reason])
  return writer.finish()
}

export function decodeTileInterestBatchV3(bytes: Uint8Array): TileInterestBatchV3 {
  const reader = new BinaryReader(bytes)
  const magic = reader.u32()
  if (magic !== TILE_INTEREST_V3_MAGIC) {
    throw new BinaryProtocolError('Invalid tile interest v3 magic')
  }
  const version = reader.u32()
  if (version !== TILE_INTEREST_V3_VERSION) {
    throw new BinaryProtocolError(`Unsupported tile interest v3 version ${version}`)
  }
  const batch: TileInterestBatchV3 = {
    magic: 'bilig.tile.interest.v3',
    version: 1,
    seq: reader.u32(),
    sheetId: reader.u32(),
    sheetOrdinal: reader.u32(),
    cameraSeq: reader.u32(),
    axisSeqX: reader.u32(),
    axisSeqY: reader.u32(),
    freezeSeq: reader.u32(),
    visibleTileKeys: decodeTileKeys(reader),
    warmTileKeys: decodeTileKeys(reader),
    pinnedTileKeys: decodeTileKeys(reader),
    reason: decodeReason(reader.u8()),
  }
  if (!reader.done()) {
    throw new BinaryProtocolError('Trailing bytes in tile interest v3 batch')
  }
  return batch
}

function encodeTileKeys(writer: BinaryWriter, keys: readonly TileKey53[]): void {
  writer.u32(keys.length)
  keys.forEach((key) => encodeTileKey(writer, key))
}

function decodeTileKeys(reader: BinaryReader): TileKey53[] {
  const count = reader.u32()
  const keys: TileKey53[] = []
  for (let index = 0; index < count; index += 1) {
    keys.push(decodeTileKey(reader))
  }
  return keys
}

function encodeTileKey(writer: BinaryWriter, key: TileKey53): void {
  if (!Number.isSafeInteger(key) || key < 0) {
    throw new BinaryProtocolError('Tile interest key must be a non-negative safe integer')
  }
  writer.f64(key)
}

function decodeTileKey(reader: BinaryReader): TileKey53 {
  const key = reader.f64()
  if (!Number.isSafeInteger(key) || key < 0) {
    throw new BinaryProtocolError('Invalid tile interest key')
  }
  return key
}

function decodeReason(tag: number): TileInterestReasonV3 {
  const reason = REASON_BY_TAG.get(tag)
  if (!reason) {
    throw new BinaryProtocolError(`Unknown tile interest reason tag ${tag}`)
  }
  return reason
}
