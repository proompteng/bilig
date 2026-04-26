import { BinaryProtocolError, BinaryReader, BinaryWriter } from '@bilig/binary-protocol'
import {
  type ErrorCode,
  ValueTag,
  type CellBorderStyle,
  type CellBorderWeight,
  type CellHorizontalAlignment,
  type CellSnapshot,
  type CellStyleRecord,
  type CellVerticalAlignment,
  type LiteralInput,
  type RecalcMetrics,
  type Viewport,
} from '@bilig/protocol'

export interface ViewportPatchSubscription extends Viewport {
  sheetName: string
  initialPatch?: 'full' | 'none'
}

export interface ViewportPatchedCell {
  row: number
  col: number
  snapshot: CellSnapshot
  displayText: string
  copyText: string
  editorText: string
  formatId: number
  styleId: string
}

export interface ViewportAxisPatch {
  index: number
  size: number
  hidden: boolean
}

export interface ViewportPatch {
  version: number
  full: boolean
  freezeRows?: number
  freezeCols?: number
  viewport: ViewportPatchSubscription
  metrics: RecalcMetrics
  styles: CellStyleRecord[]
  cells: ViewportPatchedCell[]
  columns: ViewportAxisPatch[]
  rows: ViewportAxisPatch[]
}

const VIEWPORT_PATCH_MAGIC = 0x56505450
const VIEWPORT_PATCH_CODEC_VERSION = 1
const LEGACY_JSON_OBJECT_PREFIX = 0x7b
const encoder = new TextEncoder()
const decoder = new TextDecoder()
const OPTIONAL_ABSENT = 0xff

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isViewportPatch(value: unknown): value is ViewportPatch {
  if (!isRecord(value) || !isRecord(value['viewport']) || !isRecord(value['metrics'])) {
    return false
  }
  return Array.isArray(value['styles']) && Array.isArray(value['cells']) && Array.isArray(value['columns']) && Array.isArray(value['rows'])
}

function encodeColumn(index: number): string {
  let value = index
  let result = ''
  do {
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26) - 1
  } while (value >= 0)
  return result
}

function formatAddress(row: number, col: number): string {
  return `${encodeColumn(col)}${row + 1}`
}

function encodeOptionalString(writer: BinaryWriter, value: string | undefined): void {
  writer.u8(value === undefined ? 0 : 1)
  if (value !== undefined) {
    writer.string(value)
  }
}

function decodeOptionalString(reader: BinaryReader): string | undefined {
  return reader.u8() === 0 ? undefined : reader.string()
}

function encodeOptionalNumber(writer: BinaryWriter, value: number | undefined): void {
  writer.u8(value === undefined ? 0 : 1)
  if (value !== undefined) {
    writer.f64(value)
  }
}

function decodeOptionalNumber(reader: BinaryReader): number | undefined {
  return reader.u8() === 0 ? undefined : reader.f64()
}

function encodeOptionalBoolean(writer: BinaryWriter, value: boolean | undefined): void {
  writer.u8(value === undefined ? 0 : 1)
  if (value !== undefined) {
    writer.bool(value)
  }
}

function decodeOptionalBoolean(reader: BinaryReader): boolean | undefined {
  return reader.u8() === 0 ? undefined : reader.bool()
}

function encodeOptionalTag(writer: BinaryWriter, value: number | undefined): void {
  writer.u8(value ?? OPTIONAL_ABSENT)
}

function decodeOptionalTag(reader: BinaryReader): number | undefined {
  const value = reader.u8()
  return value === OPTIONAL_ABSENT ? undefined : value
}

function encodeLiteralInput(writer: BinaryWriter, value: LiteralInput | undefined): void {
  if (value === undefined) {
    writer.u8(0)
    return
  }
  if (value === null) {
    writer.u8(1)
    return
  }
  if (typeof value === 'number') {
    writer.u8(2)
    writer.f64(value)
    return
  }
  if (typeof value === 'string') {
    writer.u8(3)
    writer.string(value)
    return
  }
  writer.u8(4)
  writer.bool(value)
}

function decodeLiteralInput(reader: BinaryReader): LiteralInput | undefined {
  switch (reader.u8()) {
    case 0:
      return undefined
    case 1:
      return null
    case 2:
      return reader.f64()
    case 3:
      return reader.string()
    case 4:
      return reader.bool()
    default:
      throw new BinaryProtocolError('Unknown literal input tag')
  }
}

function encodeHorizontalAlignment(value: CellHorizontalAlignment): number {
  switch (value) {
    case 'general':
      return 0
    case 'left':
      return 1
    case 'center':
      return 2
    case 'right':
      return 3
  }
}

function decodeHorizontalAlignment(value: number): CellHorizontalAlignment {
  switch (value) {
    case 0:
      return 'general'
    case 1:
      return 'left'
    case 2:
      return 'center'
    case 3:
      return 'right'
    default:
      throw new BinaryProtocolError(`Unknown horizontal alignment tag ${value}`)
  }
}

function encodeVerticalAlignment(value: CellVerticalAlignment): number {
  switch (value) {
    case 'top':
      return 0
    case 'middle':
      return 1
    case 'bottom':
      return 2
  }
}

function decodeVerticalAlignment(value: number): CellVerticalAlignment {
  switch (value) {
    case 0:
      return 'top'
    case 1:
      return 'middle'
    case 2:
      return 'bottom'
    default:
      throw new BinaryProtocolError(`Unknown vertical alignment tag ${value}`)
  }
}

function encodeBorderStyle(value: CellBorderStyle): number {
  switch (value) {
    case 'solid':
      return 0
    case 'dashed':
      return 1
    case 'dotted':
      return 2
    case 'double':
      return 3
  }
}

function decodeBorderStyle(value: number): CellBorderStyle {
  switch (value) {
    case 0:
      return 'solid'
    case 1:
      return 'dashed'
    case 2:
      return 'dotted'
    case 3:
      return 'double'
    default:
      throw new BinaryProtocolError(`Unknown border style tag ${value}`)
  }
}

function encodeBorderWeight(value: CellBorderWeight): number {
  switch (value) {
    case 'thin':
      return 0
    case 'medium':
      return 1
    case 'thick':
      return 2
  }
}

function decodeBorderWeight(value: number): CellBorderWeight {
  switch (value) {
    case 0:
      return 'thin'
    case 1:
      return 'medium'
    case 2:
      return 'thick'
    default:
      throw new BinaryProtocolError(`Unknown border weight tag ${value}`)
  }
}

function encodeCellValue(writer: BinaryWriter, value: CellSnapshot['value']): void {
  writer.u8(value.tag)
  switch (value.tag) {
    case ValueTag.Empty:
      return
    case ValueTag.Number:
      writer.f64(value.value)
      return
    case ValueTag.Boolean:
      writer.bool(value.value)
      return
    case ValueTag.String:
      writer.u32(value.stringId)
      writer.string(value.value)
      return
    case ValueTag.Error:
      writer.u8(value.code)
      return
    default:
      throw new BinaryProtocolError('Unknown cell value tag')
  }
}

function decodeCellValue(reader: BinaryReader): CellSnapshot['value'] {
  const tag = reader.u8() as ValueTag
  switch (tag) {
    case ValueTag.Empty:
      return { tag: ValueTag.Empty }
    case ValueTag.Number:
      return { tag: ValueTag.Number, value: reader.f64() }
    case ValueTag.Boolean:
      return { tag: ValueTag.Boolean, value: reader.bool() }
    case ValueTag.String:
      return {
        tag: ValueTag.String,
        stringId: reader.u32(),
        value: reader.string(),
      }
    case ValueTag.Error:
      return { tag: ValueTag.Error, code: reader.u8() as ErrorCode }
    default:
      throw new BinaryProtocolError(`Unknown cell value tag ${String(tag)}`)
  }
}

function encodeCellSnapshot(writer: BinaryWriter, snapshot: CellSnapshot): void {
  encodeOptionalString(writer, snapshot.formula)
  encodeOptionalString(writer, snapshot.format)
  encodeOptionalString(writer, snapshot.numberFormatId)
  encodeOptionalString(writer, snapshot.styleId)
  encodeLiteralInput(writer, snapshot.input)
  encodeCellValue(writer, snapshot.value)
  writer.u32(snapshot.flags)
  writer.u32(snapshot.version)
}

function decodeCellSnapshot(reader: BinaryReader, sheetName: string, row: number, col: number): CellSnapshot {
  const formula = decodeOptionalString(reader)
  const format = decodeOptionalString(reader)
  const numberFormatId = decodeOptionalString(reader)
  const styleId = decodeOptionalString(reader)
  const input = decodeLiteralInput(reader)
  return {
    sheetName,
    address: formatAddress(row, col),
    ...(formula !== undefined ? { formula } : {}),
    ...(format !== undefined ? { format } : {}),
    ...(numberFormatId !== undefined ? { numberFormatId } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
    ...(input !== undefined ? { input } : {}),
    value: decodeCellValue(reader),
    flags: reader.u32(),
    version: reader.u32(),
  }
}

function encodeBorderSide(writer: BinaryWriter, side: NonNullable<NonNullable<CellStyleRecord['borders']>['top']> | undefined): void {
  if (!side) {
    writer.u8(0)
    return
  }
  writer.u8(1)
  writer.u8(encodeBorderStyle(side.style))
  writer.u8(encodeBorderWeight(side.weight))
  writer.string(side.color)
}

function decodeBorderSide(reader: BinaryReader): NonNullable<NonNullable<CellStyleRecord['borders']>['top']> | undefined {
  if (reader.u8() === 0) {
    return undefined
  }
  return {
    style: decodeBorderStyle(reader.u8()),
    weight: decodeBorderWeight(reader.u8()),
    color: reader.string(),
  }
}

function encodeCellStyleRecord(writer: BinaryWriter, style: CellStyleRecord): void {
  writer.string(style.id)

  writer.u8(style.fill ? 1 : 0)
  if (style.fill) {
    writer.string(style.fill.backgroundColor)
  }

  writer.u8(style.font ? 1 : 0)
  if (style.font) {
    encodeOptionalString(writer, style.font.family)
    encodeOptionalNumber(writer, style.font.size)
    encodeOptionalBoolean(writer, style.font.bold)
    encodeOptionalBoolean(writer, style.font.italic)
    encodeOptionalBoolean(writer, style.font.underline)
    encodeOptionalString(writer, style.font.color)
  }

  writer.u8(style.alignment ? 1 : 0)
  if (style.alignment) {
    encodeOptionalTag(writer, style.alignment.horizontal ? encodeHorizontalAlignment(style.alignment.horizontal) : undefined)
    encodeOptionalTag(writer, style.alignment.vertical ? encodeVerticalAlignment(style.alignment.vertical) : undefined)
    encodeOptionalBoolean(writer, style.alignment.wrap)
    encodeOptionalNumber(writer, style.alignment.indent)
  }

  writer.u8(style.borders ? 1 : 0)
  if (style.borders) {
    encodeBorderSide(writer, style.borders.top)
    encodeBorderSide(writer, style.borders.right)
    encodeBorderSide(writer, style.borders.bottom)
    encodeBorderSide(writer, style.borders.left)
  }
}

function decodeCellStyleRecord(reader: BinaryReader): CellStyleRecord {
  const style: CellStyleRecord = { id: reader.string() }

  if (reader.u8() === 1) {
    style.fill = { backgroundColor: reader.string() }
  }

  if (reader.u8() === 1) {
    const family = decodeOptionalString(reader)
    const size = decodeOptionalNumber(reader)
    const bold = decodeOptionalBoolean(reader)
    const italic = decodeOptionalBoolean(reader)
    const underline = decodeOptionalBoolean(reader)
    const color = decodeOptionalString(reader)
    style.font = {
      ...(family !== undefined ? { family } : {}),
      ...(size !== undefined ? { size } : {}),
      ...(bold !== undefined ? { bold } : {}),
      ...(italic !== undefined ? { italic } : {}),
      ...(underline !== undefined ? { underline } : {}),
      ...(color !== undefined ? { color } : {}),
    }
  }

  if (reader.u8() === 1) {
    const horizontal = decodeOptionalTag(reader)
    const vertical = decodeOptionalTag(reader)
    const wrap = decodeOptionalBoolean(reader)
    const indent = decodeOptionalNumber(reader)
    style.alignment = {
      ...(horizontal !== undefined ? { horizontal: decodeHorizontalAlignment(horizontal) } : {}),
      ...(vertical !== undefined ? { vertical: decodeVerticalAlignment(vertical) } : {}),
      ...(wrap !== undefined ? { wrap } : {}),
      ...(indent !== undefined ? { indent } : {}),
    }
  }

  if (reader.u8() === 1) {
    const top = decodeBorderSide(reader)
    const right = decodeBorderSide(reader)
    const bottom = decodeBorderSide(reader)
    const left = decodeBorderSide(reader)
    style.borders = {
      ...(top ? { top } : {}),
      ...(right ? { right } : {}),
      ...(bottom ? { bottom } : {}),
      ...(left ? { left } : {}),
    }
  }

  return style
}

function encodeMetrics(writer: BinaryWriter, metrics: RecalcMetrics): void {
  writer.u32(metrics.batchId)
  writer.u32(metrics.changedInputCount)
  writer.u32(metrics.dirtyFormulaCount)
  writer.u32(metrics.wasmFormulaCount)
  writer.u32(metrics.jsFormulaCount)
  writer.u32(metrics.rangeNodeVisits)
  writer.f64(metrics.recalcMs)
  writer.f64(metrics.compileMs)
}

function decodeMetrics(reader: BinaryReader): RecalcMetrics {
  return {
    batchId: reader.u32(),
    changedInputCount: reader.u32(),
    dirtyFormulaCount: reader.u32(),
    wasmFormulaCount: reader.u32(),
    jsFormulaCount: reader.u32(),
    rangeNodeVisits: reader.u32(),
    recalcMs: reader.f64(),
    compileMs: reader.f64(),
  }
}

function encodeAxisPatch(writer: BinaryWriter, patch: ViewportAxisPatch): void {
  writer.u32(patch.index)
  writer.f64(patch.size)
  writer.bool(patch.hidden)
}

function decodeAxisPatch(reader: BinaryReader): ViewportAxisPatch {
  return {
    index: reader.u32(),
    size: reader.f64(),
    hidden: reader.bool(),
  }
}

function encodeAxisPatches(writer: BinaryWriter, patches: readonly ViewportAxisPatch[]): void {
  writer.u32(patches.length)
  patches.forEach((patch) => encodeAxisPatch(writer, patch))
}

function decodeAxisPatches(reader: BinaryReader): ViewportAxisPatch[] {
  const count = reader.u32()
  const patches: ViewportAxisPatch[] = []
  for (let index = 0; index < count; index += 1) {
    patches.push(decodeAxisPatch(reader))
  }
  return patches
}

function decodeLegacyJson(bytes: Uint8Array): ViewportPatch {
  const parsed = JSON.parse(decoder.decode(bytes)) as unknown
  if (!isViewportPatch(parsed)) {
    throw new Error('Invalid viewport patch payload')
  }
  return parsed
}

export function encodeViewportPatch(patch: ViewportPatch): Uint8Array {
  const writer = new BinaryWriter()
  writer.u32(VIEWPORT_PATCH_MAGIC)
  writer.u32(VIEWPORT_PATCH_CODEC_VERSION)
  writer.u32(patch.version)
  writer.bool(patch.full)
  writer.u32(patch.freezeRows ?? 0)
  writer.u32(patch.freezeCols ?? 0)
  writer.string(patch.viewport.sheetName)
  writer.u32(patch.viewport.rowStart)
  writer.u32(patch.viewport.rowEnd)
  writer.u32(patch.viewport.colStart)
  writer.u32(patch.viewport.colEnd)
  encodeMetrics(writer, patch.metrics)
  writer.u32(patch.styles.length)
  patch.styles.forEach((style) => encodeCellStyleRecord(writer, style))
  writer.u32(patch.cells.length)
  patch.cells.forEach((cell) => {
    writer.u32(cell.row)
    writer.u32(cell.col)
    encodeCellSnapshot(writer, cell.snapshot)
    writer.string(cell.displayText)
    writer.string(cell.copyText)
    writer.string(cell.editorText)
    writer.u32(cell.formatId)
    writer.string(cell.styleId)
  })
  encodeAxisPatches(writer, patch.columns)
  encodeAxisPatches(writer, patch.rows)
  return writer.finish()
}

export function decodeViewportPatch(bytes: Uint8Array): ViewportPatch {
  if (bytes.byteLength === 0) {
    throw new Error('Invalid viewport patch payload')
  }
  if (bytes[0] === LEGACY_JSON_OBJECT_PREFIX) {
    return decodeLegacyJson(bytes)
  }

  const reader = new BinaryReader(bytes)
  const magic = reader.u32()
  if (magic !== VIEWPORT_PATCH_MAGIC) {
    return decodeLegacyJson(bytes)
  }

  const codecVersion = reader.u32()
  if (codecVersion !== VIEWPORT_PATCH_CODEC_VERSION) {
    throw new BinaryProtocolError(`Unsupported viewport patch codec version ${codecVersion}`)
  }

  const version = reader.u32()
  const full = reader.bool()
  const freezeRows = reader.u32()
  const freezeCols = reader.u32()
  const viewport: ViewportPatchSubscription = {
    sheetName: reader.string(),
    rowStart: reader.u32(),
    rowEnd: reader.u32(),
    colStart: reader.u32(),
    colEnd: reader.u32(),
  }
  const patch: ViewportPatch = {
    version,
    full,
    freezeRows,
    freezeCols,
    viewport,
    metrics: decodeMetrics(reader),
    styles: [],
    cells: [],
    columns: [],
    rows: [],
  }

  const styleCount = reader.u32()
  for (let index = 0; index < styleCount; index += 1) {
    patch.styles.push(decodeCellStyleRecord(reader))
  }

  const cellCount = reader.u32()
  for (let index = 0; index < cellCount; index += 1) {
    const row = reader.u32()
    const col = reader.u32()
    patch.cells.push({
      row,
      col,
      snapshot: decodeCellSnapshot(reader, viewport.sheetName, row, col),
      displayText: reader.string(),
      copyText: reader.string(),
      editorText: reader.string(),
      formatId: reader.u32(),
      styleId: reader.string(),
    })
  }

  patch.columns = decodeAxisPatches(reader)
  patch.rows = decodeAxisPatches(reader)

  if (!reader.done()) {
    throw new BinaryProtocolError('Viewport patch payload has trailing bytes')
  }

  return patch
}

export function encodeViewportPatchJson(patch: ViewportPatch): Uint8Array {
  return encoder.encode(JSON.stringify(patch))
}
