import type {
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellNumberFormatKind,
  CellNumberFormatRecord,
  CellRangeRef,
  CellStylePatch,
  CellStyleRecord,
  CellVerticalAlignment,
  CompatibilityMode,
  LiteralInput,
  PivotAggregation,
  WorkbookAxisEntrySnapshot,
  WorkbookCalculationMode,
  WorkbookDefinedNameValueSnapshot,
} from '@bilig/protocol'
import type { WorkbookSortDirection, WorkbookSortKey } from '@bilig/workbook-domain'
import { BinaryProtocolError, type BinaryReader, type BinaryWriter } from './binary-io.js'

type LiteralTag = 0 | 1 | 2 | 3
type DefinedNameValueTag = 0 | 1 | 2 | 3 | 4 | 5

export function assertNever(value: never): never {
  throw new BinaryProtocolError(`Unsupported value: ${String(value)}`)
}

function decodeLiteralTag(tag: number): LiteralTag {
  switch (tag) {
    case 0:
    case 1:
    case 2:
    case 3:
      return tag
    default:
      throw new BinaryProtocolError(`Unknown literal tag ${tag}`)
  }
}

function decodeDefinedNameValueTag(tag: number): DefinedNameValueTag {
  switch (tag) {
    case 0:
    case 1:
    case 2:
    case 3:
    case 4:
    case 5:
      return tag
    default:
      throw new BinaryProtocolError(`Unknown defined-name value tag ${tag}`)
  }
}

export function encodeLiteral(writer: BinaryWriter, literal: LiteralInput): void {
  if (literal === null) {
    writer.u8(0 satisfies LiteralTag)
    return
  }

  if (typeof literal === 'number') {
    writer.u8(1 satisfies LiteralTag)
    writer.f64(literal)
    return
  }
  if (typeof literal === 'string') {
    writer.u8(2 satisfies LiteralTag)
    writer.string(literal)
    return
  }
  if (typeof literal === 'boolean') {
    writer.u8(3 satisfies LiteralTag)
    writer.bool(literal)
    return
  }
  throw new BinaryProtocolError(`Unsupported literal type: ${typeof literal}`)
}

export function decodeLiteral(reader: BinaryReader): LiteralInput {
  switch (decodeLiteralTag(reader.u8())) {
    case 0:
      return null
    case 1:
      return reader.f64()
    case 2:
      return reader.string()
    case 3:
      return reader.bool()
    default:
      throw new BinaryProtocolError('Unknown literal tag')
  }
}

export function encodeDefinedNameValue(writer: BinaryWriter, value: WorkbookDefinedNameValueSnapshot): void {
  if (typeof value !== 'object' || value === null || !('kind' in value)) {
    writer.u8(0 satisfies DefinedNameValueTag)
    encodeLiteral(writer, value)
    return
  }
  switch (value.kind) {
    case 'scalar':
      writer.u8(1 satisfies DefinedNameValueTag)
      encodeLiteral(writer, value.value)
      return
    case 'cell-ref':
      writer.u8(2 satisfies DefinedNameValueTag)
      writer.string(value.sheetName)
      writer.string(value.address)
      return
    case 'range-ref':
      writer.u8(3 satisfies DefinedNameValueTag)
      encodeCellRangeRef(writer, value)
      return
    case 'structured-ref':
      writer.u8(4 satisfies DefinedNameValueTag)
      writer.string(value.tableName)
      writer.string(value.columnName)
      return
    case 'formula':
      writer.u8(5 satisfies DefinedNameValueTag)
      writer.string(value.formula)
      return
  }
}

export function decodeDefinedNameValue(reader: BinaryReader): WorkbookDefinedNameValueSnapshot {
  switch (decodeDefinedNameValueTag(reader.u8())) {
    case 0:
      return decodeLiteral(reader)
    case 1:
      return { kind: 'scalar', value: decodeLiteral(reader) }
    case 2:
      return { kind: 'cell-ref', sheetName: reader.string(), address: reader.string() }
    case 3:
      return { kind: 'range-ref', ...decodeCellRangeRef(reader) }
    case 4:
      return { kind: 'structured-ref', tableName: reader.string(), columnName: reader.string() }
    case 5:
      return { kind: 'formula', formula: reader.string() }
    default:
      throw new BinaryProtocolError('Unknown defined-name value tag')
  }
}

export function encodeNullableNumber(writer: BinaryWriter, value: number | null): void {
  writer.bool(value !== null)
  if (value !== null) {
    writer.f64(value)
  }
}

export function decodeNullableNumber(reader: BinaryReader): number | null {
  return reader.bool() ? reader.f64() : null
}

export function encodeNullableBoolean(writer: BinaryWriter, value: boolean | null): void {
  writer.u8(value === null ? 0 : value ? 2 : 1)
}

export function decodeNullableBoolean(reader: BinaryReader): boolean | null {
  switch (reader.u8()) {
    case 0:
      return null
    case 1:
      return false
    case 2:
      return true
    default:
      throw new BinaryProtocolError('Unknown nullable boolean tag')
  }
}

function encodeOptionalNullableNumber(writer: BinaryWriter, value: number | null | undefined): void {
  writer.bool(value !== undefined)
  if (value !== undefined) {
    encodeNullableNumber(writer, value)
  }
}

function decodeOptionalNullableNumber(reader: BinaryReader): number | null | undefined {
  return reader.bool() ? decodeNullableNumber(reader) : undefined
}

function encodeOptionalNullableBoolean(writer: BinaryWriter, value: boolean | null | undefined): void {
  writer.bool(value !== undefined)
  if (value !== undefined) {
    encodeNullableBoolean(writer, value)
  }
}

function decodeOptionalNullableBoolean(reader: BinaryReader): boolean | null | undefined {
  return reader.bool() ? decodeNullableBoolean(reader) : undefined
}

function encodeOptionalNullableString(writer: BinaryWriter, value: string | null | undefined): void {
  writer.bool(value !== undefined)
  if (value !== undefined) {
    writer.bool(value !== null)
    if (value !== null) {
      writer.string(value)
    }
  }
}

function decodeOptionalNullableString(reader: BinaryReader): string | null | undefined {
  if (!reader.bool()) {
    return undefined
  }
  return reader.bool() ? reader.string() : null
}

export function encodeCellRangeRef(writer: BinaryWriter, ref: CellRangeRef): void {
  writer.string(ref.sheetName)
  writer.string(ref.startAddress)
  writer.string(ref.endAddress)
}

export function decodeCellRangeRef(reader: BinaryReader): CellRangeRef {
  return {
    sheetName: reader.string(),
    startAddress: reader.string(),
    endAddress: reader.string(),
  }
}

export function encodeCellStyleRecord(writer: BinaryWriter, style: CellStyleRecord): void {
  writer.string(style.id)
  writer.bool(style.fill !== undefined)
  if (style.fill) {
    writer.string(style.fill.backgroundColor)
  }
  writer.bool(style.font !== undefined)
  if (style.font) {
    writer.bool(style.font.family !== undefined)
    if (style.font.family !== undefined) {
      writer.string(style.font.family)
    }
    writer.bool(style.font.size !== undefined)
    if (style.font.size !== undefined) {
      writer.f64(style.font.size)
    }
    writer.bool(style.font.bold !== undefined)
    if (style.font.bold !== undefined) {
      writer.bool(style.font.bold)
    }
    writer.bool(style.font.italic !== undefined)
    if (style.font.italic !== undefined) {
      writer.bool(style.font.italic)
    }
    writer.bool(style.font.underline !== undefined)
    if (style.font.underline !== undefined) {
      writer.bool(style.font.underline)
    }
    writer.bool(style.font.color !== undefined)
    if (style.font.color !== undefined) {
      writer.string(style.font.color)
    }
  }
  writer.bool(style.alignment !== undefined)
  if (style.alignment) {
    writer.bool(style.alignment.horizontal !== undefined)
    if (style.alignment.horizontal !== undefined) {
      writer.string(style.alignment.horizontal)
    }
    writer.bool(style.alignment.vertical !== undefined)
    if (style.alignment.vertical !== undefined) {
      writer.string(style.alignment.vertical)
    }
    writer.bool(style.alignment.wrap !== undefined)
    if (style.alignment.wrap !== undefined) {
      writer.bool(style.alignment.wrap)
    }
    writer.bool(style.alignment.indent !== undefined)
    if (style.alignment.indent !== undefined) {
      writer.u32(style.alignment.indent)
    }
  }
  writer.bool(style.borders !== undefined)
  if (style.borders) {
    encodeBorderSide(writer, style.borders.top)
    encodeBorderSide(writer, style.borders.right)
    encodeBorderSide(writer, style.borders.bottom)
    encodeBorderSide(writer, style.borders.left)
  }
  writer.bool(style.protection !== undefined)
  if (style.protection) {
    writer.bool(style.protection.locked !== undefined)
    if (style.protection.locked !== undefined) {
      writer.bool(style.protection.locked)
    }
    writer.bool(style.protection.hidden !== undefined)
    if (style.protection.hidden !== undefined) {
      writer.bool(style.protection.hidden)
    }
  }
}

export function encodeCellNumberFormatRecord(writer: BinaryWriter, format: CellNumberFormatRecord): void {
  writer.string(format.id)
  writer.string(format.code)
  writer.string(format.kind)
}

export function decodeCellNumberFormatRecord(reader: BinaryReader): CellNumberFormatRecord {
  const id = reader.string()
  const code = reader.string()
  const kind = decodeCellNumberFormatKind(reader.string())
  return {
    id,
    code,
    kind,
  }
}

export function decodeCellStyleRecord(reader: BinaryReader): CellStyleRecord {
  const style: CellStyleRecord = { id: reader.string() }
  if (reader.bool()) {
    style.fill = { backgroundColor: reader.string() }
  }
  if (reader.bool()) {
    const font: NonNullable<CellStyleRecord['font']> = {}
    if (reader.bool()) {
      font.family = reader.string()
    }
    if (reader.bool()) {
      font.size = reader.f64()
    }
    if (reader.bool()) {
      font.bold = reader.bool()
    }
    if (reader.bool()) {
      font.italic = reader.bool()
    }
    if (reader.bool()) {
      font.underline = reader.bool()
    }
    if (reader.bool()) {
      font.color = reader.string()
    }
    style.font = font
  }
  if (reader.bool()) {
    const alignment: NonNullable<CellStyleRecord['alignment']> = {}
    if (reader.bool()) {
      const horizontal = decodeHorizontalAlignment(reader.string())
      if (horizontal !== undefined) {
        alignment.horizontal = horizontal
      }
    }
    if (reader.bool()) {
      const vertical = decodeVerticalAlignment(reader.string())
      if (vertical !== undefined) {
        alignment.vertical = vertical
      }
    }
    if (reader.bool()) {
      alignment.wrap = reader.bool()
    }
    if (reader.bool()) {
      alignment.indent = reader.u32()
    }
    style.alignment = alignment
  }
  if (reader.bool()) {
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
  if (reader.bool()) {
    const protection: NonNullable<CellStyleRecord['protection']> = {}
    if (reader.bool()) {
      protection.locked = reader.bool()
    }
    if (reader.bool()) {
      protection.hidden = reader.bool()
    }
    style.protection = protection
  }
  return style
}

export function encodeCellStylePatch(writer: BinaryWriter, patch: CellStylePatch): void {
  writer.bool(patch.fill !== undefined)
  if (patch.fill !== undefined) {
    writer.bool(patch.fill !== null)
    if (patch.fill !== null) {
      encodeOptionalNullableString(writer, patch.fill.backgroundColor)
    }
  }
  writer.bool(patch.font !== undefined)
  if (patch.font !== undefined) {
    writer.bool(patch.font !== null)
    if (patch.font !== null) {
      encodeOptionalNullableString(writer, patch.font.family)
      encodeOptionalNullableNumber(writer, patch.font.size)
      encodeOptionalNullableBoolean(writer, patch.font.bold)
      encodeOptionalNullableBoolean(writer, patch.font.italic)
      encodeOptionalNullableBoolean(writer, patch.font.underline)
      encodeOptionalNullableString(writer, patch.font.color)
    }
  }
  writer.bool(patch.alignment !== undefined)
  if (patch.alignment !== undefined) {
    writer.bool(patch.alignment !== null)
    if (patch.alignment !== null) {
      encodeOptionalNullableString(writer, patch.alignment.horizontal)
      encodeOptionalNullableString(writer, patch.alignment.vertical)
      encodeOptionalNullableBoolean(writer, patch.alignment.wrap)
      encodeOptionalNullableNumber(writer, patch.alignment.indent)
    }
  }
  writer.bool(patch.borders !== undefined)
  if (patch.borders !== undefined) {
    writer.bool(patch.borders !== null)
    if (patch.borders !== null) {
      encodePatchBorderSide(writer, patch.borders.top)
      encodePatchBorderSide(writer, patch.borders.right)
      encodePatchBorderSide(writer, patch.borders.bottom)
      encodePatchBorderSide(writer, patch.borders.left)
    }
  }
}

export function decodeCellStylePatch(reader: BinaryReader): CellStylePatch {
  const patch: CellStylePatch = {}
  if (reader.bool()) {
    if (reader.bool()) {
      const backgroundColor = decodeOptionalNullableString(reader)
      patch.fill = backgroundColor === undefined ? {} : { backgroundColor }
    } else {
      patch.fill = null
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const font: NonNullable<CellStylePatch['font']> = {}
      const family = decodeOptionalNullableString(reader)
      if (family !== undefined) {
        font.family = family
      }
      const size = decodeOptionalNullableNumber(reader)
      if (size !== undefined) {
        font.size = size
      }
      const bold = decodeOptionalNullableBoolean(reader)
      if (bold !== undefined) {
        font.bold = bold
      }
      const italic = decodeOptionalNullableBoolean(reader)
      if (italic !== undefined) {
        font.italic = italic
      }
      const underline = decodeOptionalNullableBoolean(reader)
      if (underline !== undefined) {
        font.underline = underline
      }
      const color = decodeOptionalNullableString(reader)
      if (color !== undefined) {
        font.color = color
      }
      patch.font = font
    } else {
      patch.font = null
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const alignment: NonNullable<CellStylePatch['alignment']> = {}
      const horizontal = decodeOptionalNullableString(reader)
      if (horizontal !== undefined) {
        alignment.horizontal = horizontal === null ? null : (decodeHorizontalAlignment(horizontal) ?? null)
      }
      const vertical = decodeOptionalNullableString(reader)
      if (vertical !== undefined) {
        alignment.vertical = vertical === null ? null : (decodeVerticalAlignment(vertical) ?? null)
      }
      const wrap = decodeOptionalNullableBoolean(reader)
      if (wrap !== undefined) {
        alignment.wrap = wrap
      }
      const indent = decodeOptionalNullableNumber(reader)
      if (indent !== undefined) {
        alignment.indent = indent
      }
      patch.alignment = alignment
    } else {
      patch.alignment = null
    }
  }
  if (reader.bool()) {
    if (reader.bool()) {
      const borders: NonNullable<CellStylePatch['borders']> = {}
      const top = decodePatchBorderSide(reader)
      if (top !== undefined) {
        borders.top = top
      }
      const right = decodePatchBorderSide(reader)
      if (right !== undefined) {
        borders.right = right
      }
      const bottom = decodePatchBorderSide(reader)
      if (bottom !== undefined) {
        borders.bottom = bottom
      }
      const left = decodePatchBorderSide(reader)
      if (left !== undefined) {
        borders.left = left
      }
      patch.borders = borders
    } else {
      patch.borders = null
    }
  }
  return patch
}

type EncodedBorderSide = NonNullable<NonNullable<CellStyleRecord['borders']>['top']>

function encodeBorderSide(writer: BinaryWriter, side: EncodedBorderSide | undefined): void {
  writer.bool(side !== undefined)
  if (!side) {
    return
  }
  writer.string(side.style)
  writer.string(side.weight)
  writer.string(side.color)
}

function decodeBorderSide(reader: BinaryReader): EncodedBorderSide | undefined {
  if (!reader.bool()) {
    return undefined
  }
  const style = decodeBorderStyle(reader.string())
  const weight = decodeBorderWeight(reader.string())
  return {
    style,
    weight,
    color: reader.string(),
  }
}

type EncodedPatchBorderSide = NonNullable<NonNullable<CellStylePatch['borders']>['top']>

function encodePatchBorderSide(writer: BinaryWriter, side: EncodedPatchBorderSide | null | undefined): void {
  writer.bool(side !== undefined)
  if (side === undefined) {
    return
  }
  writer.bool(side !== null)
  if (side === null) {
    return
  }
  encodeOptionalNullableString(writer, side.style)
  encodeOptionalNullableString(writer, side.weight)
  encodeOptionalNullableString(writer, side.color)
}

function decodePatchBorderSide(reader: BinaryReader): EncodedPatchBorderSide | null | undefined {
  if (!reader.bool()) {
    return undefined
  }
  if (!reader.bool()) {
    return null
  }
  const side: EncodedPatchBorderSide = {}
  const style = decodeOptionalNullableString(reader)
  if (style !== undefined) {
    side.style = style === null ? null : decodeBorderStyle(style)
  }
  const weight = decodeOptionalNullableString(reader)
  if (weight !== undefined) {
    side.weight = weight === null ? null : decodeBorderWeight(weight)
  }
  const color = decodeOptionalNullableString(reader)
  if (color !== undefined) {
    side.color = color
  }
  return side
}

function decodeCellNumberFormatKind(value: string): CellNumberFormatKind {
  switch (value) {
    case 'number':
    case 'currency':
    case 'accounting':
    case 'percent':
    case 'date':
    case 'time':
    case 'datetime':
    case 'text':
      return value
    default:
      return 'general'
  }
}

function decodeHorizontalAlignment(value: string): CellHorizontalAlignment | undefined {
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
      return value
    default:
      return undefined
  }
}

function decodeVerticalAlignment(value: string): CellVerticalAlignment | undefined {
  switch (value) {
    case 'top':
    case 'middle':
    case 'bottom':
      return value
    default:
      return undefined
  }
}

function decodeBorderStyle(value: string): CellBorderStyle {
  switch (value) {
    case 'dashed':
    case 'dotted':
    case 'double':
      return value
    default:
      return 'solid'
  }
}

function decodeBorderWeight(value: string): CellBorderWeight {
  switch (value) {
    case 'medium':
    case 'thick':
      return value
    default:
      return 'thin'
  }
}

function encodeSortDirection(writer: BinaryWriter, direction: WorkbookSortDirection): void {
  switch (direction) {
    case 'asc':
      writer.u8(1)
      return
    case 'desc':
      writer.u8(2)
      return
  }
}

export function encodeCalculationMode(writer: BinaryWriter, mode: WorkbookCalculationMode): void {
  writer.u8(mode === 'manual' ? 2 : 1)
}

export function decodeCalculationMode(reader: BinaryReader): WorkbookCalculationMode {
  return reader.u8() === 2 ? 'manual' : 'automatic'
}

export function encodeCompatibilityMode(writer: BinaryWriter, mode: CompatibilityMode): void {
  writer.u8(mode === 'odf-1.4' ? 2 : 1)
}

export function decodeCompatibilityMode(reader: BinaryReader): CompatibilityMode {
  return reader.u8() === 2 ? 'odf-1.4' : 'excel-modern'
}

export function encodeAxisEntries(writer: BinaryWriter, entries: readonly WorkbookAxisEntrySnapshot[] | undefined): void {
  writer.u32(entries?.length ?? 0)
  entries?.forEach((entry) => {
    writer.string(entry.id)
    writer.u32(entry.index)
    encodeNullableNumber(writer, entry.size ?? null)
    encodeNullableBoolean(writer, entry.hidden ?? null)
  })
}

export function decodeAxisEntries(reader: BinaryReader): WorkbookAxisEntrySnapshot[] {
  const count = reader.u32()
  const entries: WorkbookAxisEntrySnapshot[] = []
  for (let index = 0; index < count; index += 1) {
    const entry: WorkbookAxisEntrySnapshot = {
      id: reader.string(),
      index: reader.u32(),
    }
    const size = decodeNullableNumber(reader)
    const hidden = decodeNullableBoolean(reader)
    if (size !== null) {
      entry.size = size
    }
    if (hidden !== null) {
      entry.hidden = hidden
    }
    entries.push(entry)
  }
  return entries
}

function decodeSortDirection(reader: BinaryReader): WorkbookSortDirection {
  switch (reader.u8()) {
    case 1:
      return 'asc'
    case 2:
      return 'desc'
    default:
      throw new BinaryProtocolError('Unknown sort direction tag')
  }
}

export function encodeSortKey(writer: BinaryWriter, key: WorkbookSortKey): void {
  writer.string(key.keyAddress)
  encodeSortDirection(writer, key.direction)
}

export function decodeSortKey(reader: BinaryReader): WorkbookSortKey {
  return {
    keyAddress: reader.string(),
    direction: decodeSortDirection(reader),
  }
}

export function encodePivotAggregation(writer: BinaryWriter, agg: PivotAggregation): void {
  switch (agg) {
    case 'sum':
      writer.u8(1)
      return
    case 'count':
      writer.u8(2)
      return
  }
}

export function decodePivotAggregation(reader: BinaryReader): PivotAggregation {
  switch (reader.u8()) {
    case 1:
      return 'sum'
    case 2:
      return 'count'
    default:
      throw new BinaryProtocolError('Unknown pivot aggregation tag')
  }
}
