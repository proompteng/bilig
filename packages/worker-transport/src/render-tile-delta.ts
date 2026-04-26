import { BinaryProtocolError, BinaryReader, BinaryWriter } from '@bilig/binary-protocol'
import type { Viewport } from '@bilig/protocol'

export type RenderTilePaneKind =
  | 'body'
  | 'frozenTop'
  | 'frozenLeft'
  | 'frozenCorner'
  | 'columnHeaderBody'
  | 'columnHeaderFrozen'
  | 'rowHeaderBody'
  | 'rowHeaderFrozen'
  | 'dynamicOverlay'

export interface RenderTileDeltaSubscription extends Viewport {
  readonly sheetId: number
  readonly sheetName: string
  readonly cameraSeq?: number | undefined
  readonly dprBucket?: number | undefined
  readonly initialDelta?: 'full' | 'none' | undefined
}

export interface RenderTileCoord {
  readonly sheetId: number
  readonly paneKind: RenderTilePaneKind
  readonly rowTile: number
  readonly colTile: number
  readonly dprBucket: number
}

export interface RenderTileVersion {
  readonly axisX: number
  readonly axisY: number
  readonly values: number
  readonly styles: number
  readonly text: number
  readonly freeze: number
}

export interface RenderDirtySpan {
  readonly offset: number
  readonly length: number
}

export interface RenderTileDirtySpans {
  readonly rectSpans: readonly RenderDirtySpan[]
  readonly textSpans: readonly RenderDirtySpan[]
  readonly glyphSpans: readonly RenderDirtySpan[]
}

export interface RenderTileTextRun {
  readonly text: string
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly clipX: number
  readonly clipY: number
  readonly clipWidth: number
  readonly clipHeight: number
  readonly font: string
  readonly fontSize: number
  readonly color: string
  readonly underline: boolean
  readonly strike: boolean
}

export interface RenderTileReplaceMutation {
  readonly kind: 'tileReplace'
  readonly tileId: number
  readonly coord: RenderTileCoord
  readonly version: RenderTileVersion
  readonly bounds: Viewport
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly textMetrics: Float32Array
  readonly glyphRefs: Uint32Array
  readonly textRuns: readonly RenderTileTextRun[]
  readonly textCount: number
  readonly dirty: RenderTileDirtySpans
}

export interface RenderTileCellRun {
  readonly row: number
  readonly colStart: number
  readonly colEnd: number
  readonly rectSpan: RenderDirtySpan
  readonly textSpan: RenderDirtySpan
  readonly glyphSpan: RenderDirtySpan
}

export interface RenderTileCellRunMutation {
  readonly kind: 'cellRuns'
  readonly tileId: number
  readonly version: RenderTileVersion
  readonly runs: readonly RenderTileCellRun[]
}

export interface RenderTileAxisMutation {
  readonly kind: 'axis'
  readonly axis: 'row' | 'col'
  readonly changedStart: number
  readonly changedEnd: number
  readonly axisVersion: number
}

export interface RenderTileFreezeMutation {
  readonly kind: 'freeze'
  readonly freezeRows: number
  readonly freezeCols: number
  readonly freezeVersion: number
}

export interface RenderTileInvalidateMutation {
  readonly kind: 'invalidate'
  readonly tileId: number
  readonly reason: string
}

export interface RenderTileOverlayMutation {
  readonly kind: 'overlay'
  readonly overlayRevision: number
  readonly dirtyBounds: Viewport
}

export type RenderTileMutation =
  | RenderTileReplaceMutation
  | RenderTileCellRunMutation
  | RenderTileAxisMutation
  | RenderTileFreezeMutation
  | RenderTileInvalidateMutation
  | RenderTileOverlayMutation

export interface RenderTileDeltaBatch {
  readonly magic: 'bilig.render.tile.delta'
  readonly version: 1
  readonly sheetId: number
  readonly batchId: number
  readonly cameraSeq: number
  readonly mutations: readonly RenderTileMutation[]
}

const RENDER_TILE_DELTA_MAGIC = 0x52544431
const RENDER_TILE_DELTA_VERSION = 1

const PANE_KIND_TAGS: Record<RenderTilePaneKind, number> = {
  body: 0,
  frozenTop: 1,
  frozenLeft: 2,
  frozenCorner: 3,
  columnHeaderBody: 4,
  columnHeaderFrozen: 5,
  rowHeaderBody: 6,
  rowHeaderFrozen: 7,
  dynamicOverlay: 8,
}

const PANE_KIND_BY_TAG = new Map<number, RenderTilePaneKind>([
  [0, 'body'],
  [1, 'frozenTop'],
  [2, 'frozenLeft'],
  [3, 'frozenCorner'],
  [4, 'columnHeaderBody'],
  [5, 'columnHeaderFrozen'],
  [6, 'rowHeaderBody'],
  [7, 'rowHeaderFrozen'],
  [8, 'dynamicOverlay'],
])

const MUTATION_TAGS: Record<RenderTileMutation['kind'], number> = {
  tileReplace: 1,
  cellRuns: 2,
  axis: 3,
  freeze: 4,
  invalidate: 5,
  overlay: 6,
}

export function encodeRenderTileDeltaBatch(batch: RenderTileDeltaBatch): Uint8Array {
  const writer = new BinaryWriter()
  writer.u32(RENDER_TILE_DELTA_MAGIC)
  writer.u32(RENDER_TILE_DELTA_VERSION)
  writer.u32(batch.sheetId)
  writer.u32(batch.batchId)
  writer.u32(batch.cameraSeq)
  writer.u32(batch.mutations.length)
  batch.mutations.forEach((mutation) => encodeMutation(writer, mutation))
  return writer.finish()
}

export function decodeRenderTileDeltaBatch(bytes: Uint8Array): RenderTileDeltaBatch {
  const reader = new BinaryReader(bytes)
  const magic = reader.u32()
  if (magic !== RENDER_TILE_DELTA_MAGIC) {
    throw new BinaryProtocolError('Invalid render tile delta magic')
  }
  const version = reader.u32()
  if (version !== RENDER_TILE_DELTA_VERSION) {
    throw new BinaryProtocolError(`Unsupported render tile delta version ${version}`)
  }
  const batch: RenderTileDeltaBatch = {
    magic: 'bilig.render.tile.delta',
    version: 1,
    sheetId: reader.u32(),
    batchId: reader.u32(),
    cameraSeq: reader.u32(),
    mutations: decodeArray(reader, () => decodeMutation(reader)),
  }
  if (!reader.done()) {
    throw new BinaryProtocolError('Trailing bytes in render tile delta batch')
  }
  return batch
}

function encodeMutation(writer: BinaryWriter, mutation: RenderTileMutation): void {
  writer.u8(MUTATION_TAGS[mutation.kind])
  switch (mutation.kind) {
    case 'tileReplace':
      writer.u32(mutation.tileId)
      encodeTileCoord(writer, mutation.coord)
      encodeTileVersion(writer, mutation.version)
      encodeViewport(writer, mutation.bounds)
      encodeFloat32Array(writer, mutation.rectInstances)
      writer.u32(mutation.rectCount)
      encodeFloat32Array(writer, mutation.textMetrics)
      encodeUint32Array(writer, mutation.glyphRefs)
      encodeArray(writer, mutation.textRuns, encodeTextRun)
      writer.u32(mutation.textCount)
      encodeDirtySpans(writer, mutation.dirty)
      return
    case 'cellRuns':
      writer.u32(mutation.tileId)
      encodeTileVersion(writer, mutation.version)
      encodeArray(writer, mutation.runs, encodeCellRun)
      return
    case 'axis':
      writer.u8(mutation.axis === 'row' ? 0 : 1)
      writer.u32(mutation.changedStart)
      writer.u32(mutation.changedEnd)
      writer.u32(mutation.axisVersion)
      return
    case 'freeze':
      writer.u32(mutation.freezeRows)
      writer.u32(mutation.freezeCols)
      writer.u32(mutation.freezeVersion)
      return
    case 'invalidate':
      writer.u32(mutation.tileId)
      writer.string(mutation.reason)
      return
    case 'overlay':
      writer.u32(mutation.overlayRevision)
      encodeViewport(writer, mutation.dirtyBounds)
      return
  }
}

function decodeMutation(reader: BinaryReader): RenderTileMutation {
  switch (reader.u8()) {
    case 1:
      return {
        kind: 'tileReplace',
        tileId: reader.u32(),
        coord: decodeTileCoord(reader),
        version: decodeTileVersion(reader),
        bounds: decodeViewport(reader),
        rectInstances: decodeFloat32Array(reader),
        rectCount: reader.u32(),
        textMetrics: decodeFloat32Array(reader),
        glyphRefs: decodeUint32Array(reader),
        textRuns: decodeArray(reader, () => decodeTextRun(reader)),
        textCount: reader.u32(),
        dirty: decodeDirtySpans(reader),
      }
    case 2:
      return {
        kind: 'cellRuns',
        tileId: reader.u32(),
        version: decodeTileVersion(reader),
        runs: decodeArray(reader, () => decodeCellRun(reader)),
      }
    case 3: {
      const axisTag = reader.u8()
      if (axisTag !== 0 && axisTag !== 1) {
        throw new BinaryProtocolError(`Unknown render tile axis tag ${axisTag}`)
      }
      return {
        kind: 'axis',
        axis: axisTag === 0 ? 'row' : 'col',
        changedStart: reader.u32(),
        changedEnd: reader.u32(),
        axisVersion: reader.u32(),
      }
    }
    case 4:
      return {
        kind: 'freeze',
        freezeRows: reader.u32(),
        freezeCols: reader.u32(),
        freezeVersion: reader.u32(),
      }
    case 5:
      return {
        kind: 'invalidate',
        tileId: reader.u32(),
        reason: reader.string(),
      }
    case 6:
      return {
        kind: 'overlay',
        overlayRevision: reader.u32(),
        dirtyBounds: decodeViewport(reader),
      }
    default:
      throw new BinaryProtocolError('Unknown render tile mutation tag')
  }
}

function encodeTileCoord(writer: BinaryWriter, coord: RenderTileCoord): void {
  writer.u32(coord.sheetId)
  writer.u8(PANE_KIND_TAGS[coord.paneKind])
  writer.u32(coord.rowTile)
  writer.u32(coord.colTile)
  writer.u32(coord.dprBucket)
}

function decodeTileCoord(reader: BinaryReader): RenderTileCoord {
  const sheetId = reader.u32()
  const paneKindTag = reader.u8()
  const paneKind = PANE_KIND_BY_TAG.get(paneKindTag)
  if (!paneKind) {
    throw new BinaryProtocolError(`Unknown render tile pane kind tag ${paneKindTag}`)
  }
  return {
    sheetId,
    paneKind,
    rowTile: reader.u32(),
    colTile: reader.u32(),
    dprBucket: reader.u32(),
  }
}

function encodeTileVersion(writer: BinaryWriter, version: RenderTileVersion): void {
  writer.u32(version.axisX)
  writer.u32(version.axisY)
  writer.u32(version.values)
  writer.u32(version.styles)
  writer.u32(version.text)
  writer.u32(version.freeze)
}

function decodeTileVersion(reader: BinaryReader): RenderTileVersion {
  return {
    axisX: reader.u32(),
    axisY: reader.u32(),
    values: reader.u32(),
    styles: reader.u32(),
    text: reader.u32(),
    freeze: reader.u32(),
  }
}

function encodeViewport(writer: BinaryWriter, viewport: Viewport): void {
  writer.u32(viewport.rowStart)
  writer.u32(viewport.rowEnd)
  writer.u32(viewport.colStart)
  writer.u32(viewport.colEnd)
}

function decodeViewport(reader: BinaryReader): Viewport {
  return {
    rowStart: reader.u32(),
    rowEnd: reader.u32(),
    colStart: reader.u32(),
    colEnd: reader.u32(),
  }
}

function encodeDirtySpans(writer: BinaryWriter, spans: RenderTileDirtySpans): void {
  encodeArray(writer, spans.rectSpans, encodeDirtySpan)
  encodeArray(writer, spans.textSpans, encodeDirtySpan)
  encodeArray(writer, spans.glyphSpans, encodeDirtySpan)
}

function decodeDirtySpans(reader: BinaryReader): RenderTileDirtySpans {
  return {
    rectSpans: decodeArray(reader, () => decodeDirtySpan(reader)),
    textSpans: decodeArray(reader, () => decodeDirtySpan(reader)),
    glyphSpans: decodeArray(reader, () => decodeDirtySpan(reader)),
  }
}

function encodeDirtySpan(writer: BinaryWriter, span: RenderDirtySpan): void {
  writer.u32(span.offset)
  writer.u32(span.length)
}

function decodeDirtySpan(reader: BinaryReader): RenderDirtySpan {
  return {
    offset: reader.u32(),
    length: reader.u32(),
  }
}

function encodeCellRun(writer: BinaryWriter, run: RenderTileCellRun): void {
  writer.u32(run.row)
  writer.u32(run.colStart)
  writer.u32(run.colEnd)
  encodeDirtySpan(writer, run.rectSpan)
  encodeDirtySpan(writer, run.textSpan)
  encodeDirtySpan(writer, run.glyphSpan)
}

function decodeCellRun(reader: BinaryReader): RenderTileCellRun {
  return {
    row: reader.u32(),
    colStart: reader.u32(),
    colEnd: reader.u32(),
    rectSpan: decodeDirtySpan(reader),
    textSpan: decodeDirtySpan(reader),
    glyphSpan: decodeDirtySpan(reader),
  }
}

function encodeTextRun(writer: BinaryWriter, run: RenderTileTextRun): void {
  writer.string(run.text)
  writer.f64(run.x)
  writer.f64(run.y)
  writer.f64(run.width)
  writer.f64(run.height)
  writer.f64(run.clipX)
  writer.f64(run.clipY)
  writer.f64(run.clipWidth)
  writer.f64(run.clipHeight)
  writer.string(run.font)
  writer.f64(run.fontSize)
  writer.string(run.color)
  writer.bool(run.underline)
  writer.bool(run.strike)
}

function decodeTextRun(reader: BinaryReader): RenderTileTextRun {
  return {
    text: reader.string(),
    x: reader.f64(),
    y: reader.f64(),
    width: reader.f64(),
    height: reader.f64(),
    clipX: reader.f64(),
    clipY: reader.f64(),
    clipWidth: reader.f64(),
    clipHeight: reader.f64(),
    font: reader.string(),
    fontSize: reader.f64(),
    color: reader.string(),
    underline: reader.bool(),
    strike: reader.bool(),
  }
}

function encodeFloat32Array(writer: BinaryWriter, values: Float32Array): void {
  writer.u32(values.length)
  writer.bytes(new Uint8Array(values.buffer, values.byteOffset, values.byteLength))
}

function decodeFloat32Array(reader: BinaryReader): Float32Array {
  const length = reader.u32()
  const bytes = reader.bytesView()
  if (bytes.byteLength !== length * Float32Array.BYTES_PER_ELEMENT) {
    throw new BinaryProtocolError('Invalid Float32Array payload length in render tile delta')
  }
  return new Float32Array(new Uint8Array(bytes).buffer)
}

function encodeUint32Array(writer: BinaryWriter, values: Uint32Array): void {
  writer.u32(values.length)
  writer.bytes(new Uint8Array(values.buffer, values.byteOffset, values.byteLength))
}

function decodeUint32Array(reader: BinaryReader): Uint32Array {
  const length = reader.u32()
  const bytes = reader.bytesView()
  if (bytes.byteLength !== length * Uint32Array.BYTES_PER_ELEMENT) {
    throw new BinaryProtocolError('Invalid Uint32Array payload length in render tile delta')
  }
  return new Uint32Array(new Uint8Array(bytes).buffer)
}

function encodeArray<T>(writer: BinaryWriter, values: readonly T[], encodeValue: (writer: BinaryWriter, value: T) => void): void {
  writer.u32(values.length)
  values.forEach((value) => encodeValue(writer, value))
}

function decodeArray<T>(reader: BinaryReader, decodeValue: () => T): T[] {
  const count = reader.u32()
  const values: T[] = []
  for (let index = 0; index < count; index += 1) {
    values.push(decodeValue())
  }
  return values
}
