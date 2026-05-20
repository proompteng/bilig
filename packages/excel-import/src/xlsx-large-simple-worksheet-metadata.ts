import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookAutoFilterSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookSheetFormatPrSnapshot,
} from '@bilig/protocol'
import type { LargeSimpleHyperlinkRef } from './xlsx-large-simple-hyperlinks.js'
import type { PrintPageSetupSnapshot } from './xlsx-large-simple-printer-settings.js'

const maxExpandedAxisMetadataEntries = 2_048

export interface LargeSimpleWorksheetAxisMetadata {
  readonly entries: WorkbookAxisEntrySnapshot[]
  readonly metadata: WorkbookAxisMetadataSnapshot[]
}

export interface LargeSimpleWorksheetMergeRef {
  readonly startAddress: string
  readonly endAddress: string
}

export interface LargeSimpleWorksheetScannedMetadata {
  readonly columns?: LargeSimpleWorksheetAxisMetadata
  readonly conditionalFormats?: readonly WorkbookConditionalFormatSnapshot[]
  readonly conditionalFormattingXml?: readonly string[]
  readonly dataValidations?: readonly WorkbookDataValidationSnapshot[]
  readonly drawingRelationshipId?: string
  readonly filters?: readonly WorkbookAutoFilterSnapshot[]
  readonly hyperlinks?: readonly LargeSimpleHyperlinkRef[]
  readonly rows?: LargeSimpleWorksheetAxisMetadata
  readonly merges?: readonly LargeSimpleWorksheetMergeRef[]
  readonly printPageSetup?: PrintPageSetupSnapshot
  readonly sheetFormatPr?: WorkbookSheetFormatPrSnapshot
  readonly tableRelationshipIds?: readonly string[]
}

export function readLargeSimpleMergeRanges(sheetName: string, worksheetXml: string): WorkbookMergeRangeSnapshot[] {
  const mergeCellsXml =
    /<(?:[A-Za-z_][\w.-]*:)?mergeCells\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?mergeCells)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/u.exec(
      worksheetXml,
    )?.[0] ?? ''
  return readLargeSimpleMergeRefs(mergeCellsXml).map((range) => ({
    sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }))
}

export function readLargeSimpleMergeRefs(mergeCellsXml: string): LargeSimpleWorksheetMergeRef[] {
  return [...mergeCellsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?mergeCell\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const ref = readXmlAttribute(match[0], 'ref')
    if (!ref) {
      return []
    }
    const [startAddress, endAddress] = ref.split(':')
    if (!startAddress || !endAddress || startAddress === endAddress) {
      return []
    }
    return [{ startAddress, endAddress }]
  })
}

export function readLargeSimpleColumnMetadata(worksheetXml: string): LargeSimpleWorksheetAxisMetadata {
  const entries: WorkbookAxisEntrySnapshot[] = []
  const metadata: WorkbookAxisMetadataSnapshot[] = []
  const columnsXml =
    /<(?:[A-Za-z_][\w.-]*:)?cols\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?cols)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/u.exec(
      worksheetXml,
    )?.[0] ?? ''
  for (const match of columnsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?col\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)) {
    const tag = match[0]
    const min = readPositiveIntegerAttribute(tag, 'min')
    const max = readPositiveIntegerAttribute(tag, 'max') ?? min
    if (min === null || max === null || max < min) {
      continue
    }
    const width = readNumberAttribute(tag, 'width')
    const size = width !== null && width > 0 ? Math.round(width * 6) : null
    const start = min - 1
    const count = max - min + 1
    const styleIndex = readNonNegativeIntegerAttribute(tag, 'style')
    const hidden = readOptionalBooleanAttribute(tag, 'hidden')
    const customWidth = readOptionalBooleanAttribute(tag, 'customWidth')
    const customFormat = readOptionalBooleanAttribute(tag, 'customFormat')
    const bestFit = readOptionalBooleanAttribute(tag, 'bestFit')
    const outlineLevel = readNonNegativeIntegerAttribute(tag, 'outlineLevel')
    const collapsed = readOptionalBooleanAttribute(tag, 'collapsed')
    if (
      size === null &&
      width === null &&
      styleIndex === null &&
      hidden === null &&
      customWidth === null &&
      customFormat === null &&
      bestFit === null &&
      outlineLevel === null &&
      collapsed === null
    ) {
      continue
    }
    metadata.push({
      start,
      count,
      ...(size !== null ? { size } : {}),
      ...(width !== null ? { xlsxWidth: width } : {}),
      ...(styleIndex !== null ? { styleIndex } : {}),
      ...(hidden !== null ? { hidden } : {}),
      ...(customWidth !== null ? { customWidth } : {}),
      ...(customFormat !== null ? { customFormat } : {}),
      ...(bestFit !== null ? { bestFit } : {}),
      ...(outlineLevel !== null ? { outlineLevel } : {}),
      ...(collapsed !== null ? { collapsed } : {}),
    })
    if ((size !== null || hidden === true) && entries.length + count <= maxExpandedAxisMetadataEntries) {
      for (let column = start; column < start + count; column += 1) {
        entries.push({
          id: `col:${String(column)}`,
          index: column,
          ...(size !== null ? { size } : {}),
          ...(hidden === true ? { hidden: true } : {}),
        })
      }
    }
  }
  return { entries, metadata }
}

export function readLargeSimpleRowMetadata(worksheetXml: string): LargeSimpleWorksheetAxisMetadata {
  const entries: WorkbookAxisEntrySnapshot[] = []
  const metadata: WorkbookAxisMetadataSnapshot[] = []
  for (const match of worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?row\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/gu)) {
    appendLargeSimpleRowMetadataTag(entries, metadata, match[0])
  }
  return { entries, metadata }
}

export function appendLargeSimpleRowMetadataTag(
  entries: WorkbookAxisEntrySnapshot[],
  metadata: WorkbookAxisMetadataSnapshot[],
  tag: string,
): void {
  const row = readPositiveIntegerAttribute(tag, 'r')
  if (row === null) {
    return
  }
  appendRowMetadata(entries, metadata, row - 1, {
    height: readNumberAttribute(tag, 'ht'),
    styleIndex: readNonNegativeIntegerAttribute(tag, 's'),
    hidden: readOptionalBooleanAttribute(tag, 'hidden'),
    customFormat: readOptionalBooleanAttribute(tag, 'customFormat'),
    customHeight: readOptionalBooleanAttribute(tag, 'customHeight'),
    outlineLevel: readNonNegativeIntegerAttribute(tag, 'outlineLevel'),
    collapsed: readOptionalBooleanAttribute(tag, 'collapsed'),
    thickTop: readOptionalBooleanAttribute(tag, 'thickTop'),
    thickBottom: readOptionalBooleanAttribute(tag, 'thickBottom'),
  })
}

export function readLargeSimpleSheetFormatPr(worksheetXml: string): WorkbookSheetFormatPrSnapshot | undefined {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?sheetFormatPr\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(worksheetXml)?.[0]
  return tag ? readLargeSimpleSheetFormatPrTag(tag) : undefined
}

export function appendRowMetadata(
  entries: WorkbookAxisEntrySnapshot[],
  metadata: WorkbookAxisMetadataSnapshot[],
  index: number,
  input: {
    readonly height: number | null
    readonly styleIndex: number | null
    readonly hidden: boolean | null
    readonly customFormat: boolean | null
    readonly customHeight: boolean | null
    readonly outlineLevel: number | null
    readonly collapsed: boolean | null
    readonly thickTop: boolean | null
    readonly thickBottom: boolean | null
  },
): void {
  const size = input.height !== null && input.height > 0 ? Math.round((input.height * 96) / 72) : null
  if (
    size === null &&
    input.height === null &&
    input.styleIndex === null &&
    input.hidden === null &&
    input.customFormat === null &&
    input.customHeight === null &&
    input.outlineLevel === null &&
    input.collapsed === null &&
    input.thickTop === null &&
    input.thickBottom === null
  ) {
    return
  }
  metadata.push({
    start: index,
    count: 1,
    ...(size !== null ? { size } : {}),
    ...(input.height !== null ? { xlsxHeight: input.height } : {}),
    ...(input.styleIndex !== null ? { styleIndex: input.styleIndex } : {}),
    ...(input.hidden !== null ? { hidden: input.hidden } : {}),
    ...(input.customFormat !== null ? { customFormat: input.customFormat } : {}),
    ...(input.customHeight !== null ? { customHeight: input.customHeight } : {}),
    ...(input.outlineLevel !== null ? { outlineLevel: input.outlineLevel } : {}),
    ...(input.collapsed !== null ? { collapsed: input.collapsed } : {}),
    ...(input.thickTop !== null ? { thickTop: input.thickTop } : {}),
    ...(input.thickBottom !== null ? { thickBottom: input.thickBottom } : {}),
  })
  if ((size !== null || input.hidden === true) && entries.length < maxExpandedAxisMetadataEntries) {
    entries.push({
      id: `row:${String(index)}`,
      index,
      ...(size !== null ? { size } : {}),
      ...(input.hidden === true ? { hidden: true } : {}),
    })
  }
}

export function appendColumnMetadata(
  entries: WorkbookAxisEntrySnapshot[],
  metadata: WorkbookAxisMetadataSnapshot[],
  min: number,
  max: number,
  input: {
    readonly width: number | null
    readonly styleIndex: number | null
    readonly hidden: boolean | null
    readonly customWidth: boolean | null
    readonly customFormat: boolean | null
    readonly bestFit: boolean | null
    readonly outlineLevel: number | null
    readonly collapsed: boolean | null
  },
): void {
  if (max < min) {
    return
  }
  const size = input.width !== null && input.width > 0 ? Math.round(input.width * 6) : null
  const start = min - 1
  const count = max - min + 1
  if (
    size === null &&
    input.width === null &&
    input.styleIndex === null &&
    input.hidden === null &&
    input.customWidth === null &&
    input.customFormat === null &&
    input.bestFit === null &&
    input.outlineLevel === null &&
    input.collapsed === null
  ) {
    return
  }
  metadata.push({
    start,
    count,
    ...(size !== null ? { size } : {}),
    ...(input.width !== null ? { xlsxWidth: input.width } : {}),
    ...(input.styleIndex !== null ? { styleIndex: input.styleIndex } : {}),
    ...(input.hidden !== null ? { hidden: input.hidden } : {}),
    ...(input.customWidth !== null ? { customWidth: input.customWidth } : {}),
    ...(input.customFormat !== null ? { customFormat: input.customFormat } : {}),
    ...(input.bestFit !== null ? { bestFit: input.bestFit } : {}),
    ...(input.outlineLevel !== null ? { outlineLevel: input.outlineLevel } : {}),
    ...(input.collapsed !== null ? { collapsed: input.collapsed } : {}),
  })
  if ((size !== null || input.hidden === true) && entries.length + count <= maxExpandedAxisMetadataEntries) {
    for (let column = start; column < start + count; column += 1) {
      entries.push({
        id: `col:${String(column)}`,
        index: column,
        ...(size !== null ? { size } : {}),
        ...(input.hidden === true ? { hidden: true } : {}),
      })
    }
  }
}

export function readLargeSimpleSheetFormatPrTag(tag: string): WorkbookSheetFormatPrSnapshot | undefined {
  const output: WorkbookSheetFormatPrSnapshot = {
    ...(readNumberAttribute(tag, 'baseColWidth') !== null ? { baseColWidth: readNumberAttribute(tag, 'baseColWidth') } : {}),
    ...(readNumberAttribute(tag, 'defaultColWidth') !== null ? { defaultColWidth: readNumberAttribute(tag, 'defaultColWidth') } : {}),
    ...(readNumberAttribute(tag, 'defaultRowHeight') !== null ? { defaultRowHeight: readNumberAttribute(tag, 'defaultRowHeight') } : {}),
    ...(readOptionalBooleanAttribute(tag, 'customHeight') !== null
      ? { customHeight: readOptionalBooleanAttribute(tag, 'customHeight') }
      : {}),
    ...(readNonNegativeIntegerAttribute(tag, 'outlineLevelRow') !== null
      ? { outlineLevelRow: readNonNegativeIntegerAttribute(tag, 'outlineLevelRow') }
      : {}),
    ...(readNonNegativeIntegerAttribute(tag, 'outlineLevelCol') !== null
      ? { outlineLevelCol: readNonNegativeIntegerAttribute(tag, 'outlineLevelCol') }
      : {}),
    ...(readOptionalBooleanAttribute(tag, 'thickTop') !== null ? { thickTop: readOptionalBooleanAttribute(tag, 'thickTop') } : {}),
    ...(readOptionalBooleanAttribute(tag, 'thickBottom') !== null ? { thickBottom: readOptionalBooleanAttribute(tag, 'thickBottom') } : {}),
  }
  return Object.keys(output).length > 0 ? output : undefined
}

export function readLargeSimpleDrawingRelationshipIdTag(tag: string): string | undefined {
  return readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id') ?? undefined
}

export function readLargeSimpleDrawingRelationshipId(worksheetXml: string): string | undefined {
  const tag = /<(?:[A-Za-z_][\w.-]*:)?drawing\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(worksheetXml)?.[0]
  return tag ? readLargeSimpleDrawingRelationshipIdTag(tag) : undefined
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readNumberAttribute(xml: string, attributeName: string): number | null {
  const raw = readXmlAttribute(xml, attributeName)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

function readPositiveIntegerAttribute(xml: string, attributeName: string): number | null {
  const value = readNumberAttribute(xml, attributeName)
  return Number.isInteger(value) && value !== null && value > 0 ? value : null
}

function readNonNegativeIntegerAttribute(xml: string, attributeName: string): number | null {
  const value = readNumberAttribute(xml, attributeName)
  return Number.isInteger(value) && value !== null && value >= 0 ? value : null
}

function readOptionalBooleanAttribute(xml: string, attributeName: string): boolean | null {
  const raw = readXmlAttribute(xml, attributeName)
  if (raw === null) {
    return null
  }
  if (raw === '1' || raw === 'true') {
    return true
  }
  if (raw === '0' || raw === 'false') {
    return false
  }
  return null
}
