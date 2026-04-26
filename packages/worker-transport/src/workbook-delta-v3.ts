import { BinaryProtocolError, BinaryReader, BinaryWriter } from '@bilig/binary-protocol'

export type WorkbookDeltaSourceV3 = 'localOptimistic' | 'workerAuthoritative' | 'remote' | 'load'

export interface DirtyRangesV3 {
  readonly cellRanges: Uint32Array
  readonly axisX: Uint32Array
  readonly axisY: Uint32Array
  readonly sheets?: Uint32Array | undefined
}

export interface WorkbookDeltaBatchV3 {
  readonly magic: 'bilig.workbook.delta.v3'
  readonly version: 1
  readonly seq: number
  readonly source: WorkbookDeltaSourceV3
  readonly sheetId: number
  readonly sheetOrdinal: number
  readonly valueSeq: number
  readonly styleSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
  readonly calcSeq: number
  readonly dirty: DirtyRangesV3
}

const WORKBOOK_DELTA_V3_MAGIC = 0x42445633
const WORKBOOK_DELTA_V3_VERSION = 1

const SOURCE_TAGS: Record<WorkbookDeltaSourceV3, number> = {
  localOptimistic: 1,
  workerAuthoritative: 2,
  remote: 3,
  load: 4,
}

const SOURCE_BY_TAG = new Map<number, WorkbookDeltaSourceV3>([
  [1, 'localOptimistic'],
  [2, 'workerAuthoritative'],
  [3, 'remote'],
  [4, 'load'],
])

export function encodeWorkbookDeltaBatchV3(batch: WorkbookDeltaBatchV3): Uint8Array {
  const writer = new BinaryWriter()
  writer.u32(WORKBOOK_DELTA_V3_MAGIC)
  writer.u32(WORKBOOK_DELTA_V3_VERSION)
  writer.u32(batch.seq)
  writer.u8(SOURCE_TAGS[batch.source])
  writer.u32(batch.sheetId)
  writer.u32(batch.sheetOrdinal)
  writer.u32(batch.valueSeq)
  writer.u32(batch.styleSeq)
  writer.u32(batch.axisSeqX)
  writer.u32(batch.axisSeqY)
  writer.u32(batch.freezeSeq)
  writer.u32(batch.calcSeq)
  encodeDirtyRanges(writer, batch.dirty)
  return writer.finish()
}

export function decodeWorkbookDeltaBatchV3(bytes: Uint8Array): WorkbookDeltaBatchV3 {
  const reader = new BinaryReader(bytes)
  const magic = reader.u32()
  if (magic !== WORKBOOK_DELTA_V3_MAGIC) {
    throw new BinaryProtocolError('Invalid workbook delta v3 magic')
  }
  const version = reader.u32()
  if (version !== WORKBOOK_DELTA_V3_VERSION) {
    throw new BinaryProtocolError(`Unsupported workbook delta v3 version ${version}`)
  }
  const batch: WorkbookDeltaBatchV3 = {
    magic: 'bilig.workbook.delta.v3',
    version: 1,
    seq: reader.u32(),
    source: decodeSource(reader.u8()),
    sheetId: reader.u32(),
    sheetOrdinal: reader.u32(),
    valueSeq: reader.u32(),
    styleSeq: reader.u32(),
    axisSeqX: reader.u32(),
    axisSeqY: reader.u32(),
    freezeSeq: reader.u32(),
    calcSeq: reader.u32(),
    dirty: decodeDirtyRanges(reader),
  }
  if (!reader.done()) {
    throw new BinaryProtocolError('Trailing bytes in workbook delta v3 batch')
  }
  return batch
}

function encodeDirtyRanges(writer: BinaryWriter, dirty: DirtyRangesV3): void {
  encodeUint32Array(writer, dirty.cellRanges)
  encodeUint32Array(writer, dirty.axisX)
  encodeUint32Array(writer, dirty.axisY)
  writer.bool(dirty.sheets !== undefined)
  if (dirty.sheets !== undefined) {
    encodeUint32Array(writer, dirty.sheets)
  }
}

function decodeDirtyRanges(reader: BinaryReader): DirtyRangesV3 {
  const cellRanges = decodeUint32Array(reader)
  const axisX = decodeUint32Array(reader)
  const axisY = decodeUint32Array(reader)
  const sheets = reader.bool() ? decodeUint32Array(reader) : undefined
  return {
    axisX,
    axisY,
    cellRanges,
    ...(sheets ? { sheets } : {}),
  }
}

function encodeUint32Array(writer: BinaryWriter, values: Uint32Array): void {
  writer.u32(values.length)
  for (const value of values) {
    writer.u32(value)
  }
}

function decodeUint32Array(reader: BinaryReader): Uint32Array {
  const count = reader.u32()
  const values = new Uint32Array(count)
  for (let index = 0; index < count; index += 1) {
    values[index] = reader.u32()
  }
  return values
}

function decodeSource(tag: number): WorkbookDeltaSourceV3 {
  const source = SOURCE_BY_TAG.get(tag)
  if (!source) {
    throw new BinaryProtocolError(`Unknown workbook delta v3 source tag ${tag}`)
  }
  return source
}
