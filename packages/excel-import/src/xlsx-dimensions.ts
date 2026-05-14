import { unzipSync, zipSync } from 'fflate'

import type {
  SheetMetadataSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookSheetFormatPrSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { escapeXmlAttribute, getZipText, setXmlAttribute, setZipText } from './xlsx-export-xml.js'

interface ExportRowMetadata {
  readonly rowNumber: number
  readonly size?: number
  readonly hidden?: boolean
  readonly styleIndex?: number
  readonly xlsxHeight?: number
  readonly customFormat?: boolean
  readonly customHeight?: boolean
  readonly outlineLevel?: number
  readonly collapsed?: boolean
  readonly thickTop?: boolean
  readonly thickBottom?: boolean
  readonly exact: boolean
}

interface ExportColumnMetadata {
  start: number
  count: number
  styleIndex?: number
  xlsxWidth?: number
  customFormat?: boolean
  customWidth?: boolean
  bestFit?: boolean
  hidden?: boolean
  outlineLevel?: number
  collapsed?: boolean
}

const worksheetRowElementPattern = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/gu
const worksheetRowOpeningTagPattern = /^<row\b[^>]*\/>|^<row\b[^>]*>/u

function finitePositiveNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : undefined
}

function finiteNonNegativeInteger(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0 ? value : undefined
}

function optionalBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined
}

function formatXmlNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : String(Number(value.toFixed(12)))
}

function formatXmlBoolean(value: boolean): string {
  return value ? '1' : '0'
}

function xmlAttribute(name: string, value: string | undefined): string {
  return value === undefined ? '' : ` ${name}="${escapeXmlAttribute(value)}"`
}

function removeXmlAttribute(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}="[^"]*"`, 'u'), '')
}

function readXmlAttribute(tag: string, name: string): string | null {
  const doubleQuoted = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(tag)
  if (doubleQuoted) {
    return doubleQuoted[1] ?? null
  }
  const singleQuoted = new RegExp(`\\b${name}='([^']*)'`, 'u').exec(tag)
  return singleQuoted?.[1] ?? null
}

function readXmlNumberAttribute(tag: string, name: string): number | null {
  const raw = readXmlAttribute(tag, name)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const number = Number(raw)
  return Number.isFinite(number) ? number : null
}

function readXmlPositiveIntegerAttribute(tag: string, name: string): number | null {
  const number = readXmlNumberAttribute(tag, name)
  return number !== null && Number.isSafeInteger(number) && number > 0 ? number : null
}

function readXmlOptionalBooleanAttribute(tag: string, name: string): boolean | undefined {
  const raw = readXmlAttribute(tag, name)
  if (raw === null) {
    return undefined
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function hasExactRowGeometry(row: WorkbookAxisMetadataSnapshot): boolean {
  return (
    row.styleIndex !== undefined ||
    row.xlsxHeight !== undefined ||
    row.customFormat !== undefined ||
    row.customHeight !== undefined ||
    row.outlineLevel !== undefined ||
    row.collapsed !== undefined ||
    row.thickTop !== undefined ||
    row.thickBottom !== undefined
  )
}

function hasExactColumnGeometry(column: WorkbookAxisMetadataSnapshot): boolean {
  return (
    column.styleIndex !== undefined ||
    column.xlsxWidth !== undefined ||
    column.customFormat !== undefined ||
    column.customWidth !== undefined ||
    column.bestFit !== undefined ||
    column.outlineLevel !== undefined ||
    column.collapsed !== undefined
  )
}

function expandExportRowMetadataRecord(row: WorkbookAxisMetadataSnapshot): ExportRowMetadata[] {
  if (
    !Number.isSafeInteger(row.start) ||
    row.start < 0 ||
    !Number.isSafeInteger(row.count) ||
    row.count <= 0 ||
    !hasExactRowGeometry(row)
  ) {
    return []
  }
  const xlsxHeight = finitePositiveNumber(row.xlsxHeight ?? undefined)
  const size = finitePositiveNumber(row.size ?? undefined)
  const hidden = optionalBoolean(row.hidden)
  const styleIndex = finiteNonNegativeInteger(row.styleIndex ?? undefined)
  const customFormat = optionalBoolean(row.customFormat)
  const customHeight = optionalBoolean(row.customHeight)
  const outlineLevel = finiteNonNegativeInteger(row.outlineLevel ?? undefined)
  const collapsed = optionalBoolean(row.collapsed)
  const thickTop = optionalBoolean(row.thickTop)
  const thickBottom = optionalBoolean(row.thickBottom)
  return Array.from({ length: row.count }, (_entry, offset) => ({
    rowNumber: row.start + offset + 1,
    ...(size !== undefined ? { size } : {}),
    ...(hidden !== undefined ? { hidden } : {}),
    ...(styleIndex !== undefined ? { styleIndex } : {}),
    ...(xlsxHeight !== undefined ? { xlsxHeight } : {}),
    ...(customFormat !== undefined ? { customFormat } : {}),
    ...(customHeight !== undefined ? { customHeight } : {}),
    ...(outlineLevel !== undefined ? { outlineLevel } : {}),
    ...(collapsed !== undefined ? { collapsed } : {}),
    ...(thickTop !== undefined ? { thickTop } : {}),
    ...(thickBottom !== undefined ? { thickBottom } : {}),
    exact: true,
  }))
}

function normalizeExportRowMetadata(
  rows: readonly WorkbookAxisEntrySnapshot[] | undefined,
  rowMetadata: readonly WorkbookAxisMetadataSnapshot[] | undefined,
): ExportRowMetadata[] {
  const exactRows =
    rowMetadata?.flatMap((row) => expandExportRowMetadataRecord(row)).toSorted((left, right) => left.rowNumber - right.rowNumber) ?? []
  if (!rows || rows.length === 0) {
    return exactRows
  }
  const exactRowNumbers = new Set(exactRows.map((row) => row.rowNumber))
  const fallbackRows = rows.flatMap((row) => {
    if (!Number.isSafeInteger(row.index) || row.index < 0) {
      return []
    }
    const rowNumber = row.index + 1
    if (exactRowNumbers.has(rowNumber)) {
      return []
    }
    const size = finitePositiveNumber(row.size ?? undefined)
    if (size === undefined && row.hidden !== true) {
      return []
    }
    return [
      {
        rowNumber,
        ...(size !== undefined ? { size } : {}),
        ...(row.hidden === true ? { hidden: true } : {}),
        exact: false,
      },
    ]
  })
  return [...exactRows, ...fallbackRows].toSorted((left, right) => left.rowNumber - right.rowNumber)
}

function normalizeExportColumnMetadata(columnMetadata: readonly WorkbookAxisMetadataSnapshot[] | undefined): ExportColumnMetadata[] {
  if (!columnMetadata || columnMetadata.length === 0) {
    return []
  }
  return columnMetadata
    .flatMap((column) => {
      if (
        !Number.isSafeInteger(column.start) ||
        column.start < 0 ||
        !Number.isSafeInteger(column.count) ||
        column.count <= 0 ||
        !hasExactColumnGeometry(column)
      ) {
        return []
      }
      const xlsxWidth = finitePositiveNumber(column.xlsxWidth ?? undefined)
      const styleIndex = finiteNonNegativeInteger(column.styleIndex ?? undefined)
      const customFormat = optionalBoolean(column.customFormat)
      if (xlsxWidth === undefined && styleIndex === undefined && customFormat === undefined) {
        return []
      }
      const customWidth = optionalBoolean(column.customWidth)
      const bestFit = optionalBoolean(column.bestFit)
      const hidden = optionalBoolean(column.hidden)
      const outlineLevel = finiteNonNegativeInteger(column.outlineLevel ?? undefined)
      const collapsed = optionalBoolean(column.collapsed)
      const normalized: ExportColumnMetadata = {
        start: column.start,
        count: column.count,
        ...(styleIndex !== undefined ? { styleIndex } : {}),
        ...(xlsxWidth !== undefined ? { xlsxWidth } : {}),
        ...(customFormat !== undefined ? { customFormat } : {}),
        ...(customWidth !== undefined ? { customWidth } : {}),
        ...(bestFit !== undefined ? { bestFit } : {}),
        ...(hidden !== undefined ? { hidden } : {}),
        ...(outlineLevel !== undefined ? { outlineLevel } : {}),
        ...(collapsed !== undefined ? { collapsed } : {}),
      }
      return [normalized]
    })
    .toSorted((left, right) => left.start - right.start || left.count - right.count)
}

function parseExistingColumnMetadata(sheetXml: string): ExportColumnMetadata[] {
  return [...sheetXml.matchAll(/<col\b[^>]*\/?>/gu)].flatMap((match) => {
    const columnTag = match[0]
    const min = readXmlPositiveIntegerAttribute(columnTag, 'min')
    const max = readXmlPositiveIntegerAttribute(columnTag, 'max') ?? min
    if (min === null || max === null || max < min) {
      return []
    }
    const xlsxWidth = finitePositiveNumber(readXmlNumberAttribute(columnTag, 'width') ?? undefined)
    const styleIndex = finiteNonNegativeInteger(readXmlNumberAttribute(columnTag, 'style') ?? undefined)
    const customFormat = readXmlOptionalBooleanAttribute(columnTag, 'customFormat')
    const customWidth = readXmlOptionalBooleanAttribute(columnTag, 'customWidth')
    const bestFit = readXmlOptionalBooleanAttribute(columnTag, 'bestFit')
    const hidden = readXmlOptionalBooleanAttribute(columnTag, 'hidden')
    const outlineLevel = finiteNonNegativeInteger(readXmlNumberAttribute(columnTag, 'outlineLevel') ?? undefined)
    const collapsed = readXmlOptionalBooleanAttribute(columnTag, 'collapsed')
    if (
      xlsxWidth === undefined &&
      styleIndex === undefined &&
      customFormat === undefined &&
      customWidth === undefined &&
      bestFit === undefined &&
      hidden === undefined &&
      outlineLevel === undefined &&
      collapsed === undefined
    ) {
      return []
    }
    const column: ExportColumnMetadata = {
      start: min - 1,
      count: max - min + 1,
      ...(styleIndex !== undefined ? { styleIndex } : {}),
      ...(xlsxWidth !== undefined ? { xlsxWidth } : {}),
      ...(customFormat !== undefined ? { customFormat } : {}),
      ...(customWidth !== undefined ? { customWidth } : {}),
      ...(bestFit !== undefined ? { bestFit } : {}),
      ...(hidden !== undefined ? { hidden } : {}),
      ...(outlineLevel !== undefined ? { outlineLevel } : {}),
      ...(collapsed !== undefined ? { collapsed } : {}),
    }
    return [column]
  })
}

function columnRangeEnd(column: Pick<ExportColumnMetadata, 'start' | 'count'>): number {
  return column.start + column.count - 1
}

function subtractColumnRanges(column: ExportColumnMetadata, exactColumns: readonly ExportColumnMetadata[]): ExportColumnMetadata[] {
  let segments: Array<{ start: number; count: number }> = [{ start: column.start, count: column.count }]
  for (const exactColumn of exactColumns) {
    const exactStart = exactColumn.start
    const exactEnd = columnRangeEnd(exactColumn)
    segments = segments.flatMap((segment) => {
      const segmentEnd = columnRangeEnd(segment)
      if (segmentEnd < exactStart || segment.start > exactEnd) {
        return [segment]
      }
      const output: Array<{ start: number; count: number }> = []
      if (segment.start < exactStart) {
        output.push({ start: segment.start, count: exactStart - segment.start })
      }
      if (segmentEnd > exactEnd) {
        output.push({ start: exactEnd + 1, count: segmentEnd - exactEnd })
      }
      return output
    })
  }
  return segments.map((segment) => {
    const output: ExportColumnMetadata = {
      start: segment.start,
      count: segment.count,
    }
    if (column.styleIndex !== undefined) {
      output.styleIndex = column.styleIndex
    }
    if (column.xlsxWidth !== undefined) {
      output.xlsxWidth = column.xlsxWidth
    }
    if (column.customFormat !== undefined) {
      output.customFormat = column.customFormat
    }
    if (column.customWidth !== undefined) {
      output.customWidth = column.customWidth
    }
    if (column.bestFit !== undefined) {
      output.bestFit = column.bestFit
    }
    if (column.hidden !== undefined) {
      output.hidden = column.hidden
    }
    if (column.outlineLevel !== undefined) {
      output.outlineLevel = column.outlineLevel
    }
    if (column.collapsed !== undefined) {
      output.collapsed = column.collapsed
    }
    return output
  })
}

function buildSheetFormatPrXml(sheetFormatPr: WorkbookSheetFormatPrSnapshot | undefined): string | null {
  if (!sheetFormatPr) {
    return null
  }
  const baseColWidth = finiteNonNegativeInteger(sheetFormatPr.baseColWidth ?? undefined)
  const defaultColWidth = finitePositiveNumber(sheetFormatPr.defaultColWidth ?? undefined)
  const defaultRowHeight = finitePositiveNumber(sheetFormatPr.defaultRowHeight ?? undefined)
  const customHeight = optionalBoolean(sheetFormatPr.customHeight)
  const outlineLevelRow = finiteNonNegativeInteger(sheetFormatPr.outlineLevelRow ?? undefined)
  const outlineLevelCol = finiteNonNegativeInteger(sheetFormatPr.outlineLevelCol ?? undefined)
  const thickTop = optionalBoolean(sheetFormatPr.thickTop)
  const thickBottom = optionalBoolean(sheetFormatPr.thickBottom)
  const attributes = [
    xmlAttribute('baseColWidth', baseColWidth !== undefined ? formatXmlNumber(baseColWidth) : undefined),
    xmlAttribute('defaultColWidth', defaultColWidth !== undefined ? formatXmlNumber(defaultColWidth) : undefined),
    xmlAttribute('defaultRowHeight', defaultRowHeight !== undefined ? formatXmlNumber(defaultRowHeight) : undefined),
    xmlAttribute('customHeight', customHeight !== undefined ? formatXmlBoolean(customHeight) : undefined),
    xmlAttribute('outlineLevelRow', outlineLevelRow !== undefined ? formatXmlNumber(outlineLevelRow) : undefined),
    xmlAttribute('outlineLevelCol', outlineLevelCol !== undefined ? formatXmlNumber(outlineLevelCol) : undefined),
    xmlAttribute('thickTop', thickTop !== undefined ? formatXmlBoolean(thickTop) : undefined),
    xmlAttribute('thickBottom', thickBottom !== undefined ? formatXmlBoolean(thickBottom) : undefined),
  ].join('')
  return attributes.length > 0 ? `<sheetFormatPr${attributes}/>` : null
}

function applySheetFormatPr(sheetXml: string, sheetFormatPr: WorkbookSheetFormatPrSnapshot | undefined): string {
  const sheetFormatPrXml = buildSheetFormatPrXml(sheetFormatPr)
  if (!sheetFormatPrXml) {
    return sheetXml
  }
  const existingPattern = /<sheetFormatPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetFormatPr>)/u
  if (existingPattern.test(sheetXml)) {
    return sheetXml.replace(existingPattern, sheetFormatPrXml)
  }
  const insertPattern = /<cols\b|<sheetData\b|<\/worksheet>/u
  const match = insertPattern.exec(sheetXml)
  return match ? `${sheetXml.slice(0, match.index)}${sheetFormatPrXml}${sheetXml.slice(match.index)}` : sheetXml
}

function buildColumnsXml(columns: readonly ExportColumnMetadata[]): string {
  const columnXml = columns
    .map((column) => {
      const min = column.start + 1
      const max = column.start + column.count
      return [
        '<col',
        xmlAttribute('min', String(min)),
        xmlAttribute('max', String(max)),
        xmlAttribute('style', column.styleIndex !== undefined ? formatXmlNumber(column.styleIndex) : undefined),
        xmlAttribute('width', column.xlsxWidth !== undefined ? formatXmlNumber(column.xlsxWidth) : undefined),
        xmlAttribute('customFormat', column.customFormat !== undefined ? formatXmlBoolean(column.customFormat) : undefined),
        xmlAttribute('customWidth', column.customWidth !== undefined ? formatXmlBoolean(column.customWidth) : undefined),
        xmlAttribute('bestFit', column.bestFit !== undefined ? formatXmlBoolean(column.bestFit) : undefined),
        xmlAttribute('hidden', column.hidden !== undefined ? formatXmlBoolean(column.hidden) : undefined),
        xmlAttribute('outlineLevel', column.outlineLevel !== undefined ? formatXmlNumber(column.outlineLevel) : undefined),
        xmlAttribute('collapsed', column.collapsed !== undefined ? formatXmlBoolean(column.collapsed) : undefined),
        '/>',
      ].join('')
    })
    .join('')
  return `<cols>${columnXml}</cols>`
}

function applyColumnMetadata(sheetXml: string, columns: readonly ExportColumnMetadata[]): string {
  if (columns.length === 0) {
    return sheetXml
  }
  const existingColumns = parseExistingColumnMetadata(sheetXml).flatMap((column) => subtractColumnRanges(column, columns))
  const columnsXml = buildColumnsXml(
    [...columns, ...existingColumns].toSorted((left, right) => left.start - right.start || left.count - right.count),
  )
  const existingPattern = /<cols\b[^>]*(?:\/>|>[\s\S]*?<\/cols>)/u
  if (existingPattern.test(sheetXml)) {
    return sheetXml.replace(existingPattern, columnsXml)
  }
  const match = /<sheetData\b|<\/worksheet>/u.exec(sheetXml)
  return match ? `${sheetXml.slice(0, match.index)}${columnsXml}${sheetXml.slice(match.index)}` : sheetXml
}

function readRowNumber(rowTag: string): number | null {
  const match = /\br="([0-9]+)"/u.exec(rowTag)
  if (!match) {
    return null
  }
  const rowNumber = Number(match[1])
  return Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber : null
}

function clearManagedRowAttributes(rowTag: string): string {
  return ['s', 'customFormat', 'ht', 'customHeight', 'hidden', 'outlineLevel', 'collapsed', 'thickTop', 'thickBot'].reduce(
    (tag, attribute) => removeXmlAttribute(tag, attribute),
    rowTag,
  )
}

function applyRowMetadata(rowTag: string, row: ExportRowMetadata): string {
  let nextTag = clearManagedRowAttributes(rowTag)
  const height = row.xlsxHeight ?? row.size
  if (row.styleIndex !== undefined) {
    nextTag = setXmlAttribute(nextTag, 's', formatXmlNumber(row.styleIndex))
  }
  if (row.customFormat !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'customFormat', formatXmlBoolean(row.customFormat))
  }
  if (height !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'ht', formatXmlNumber(height))
    if (!row.exact || row.customHeight !== undefined) {
      nextTag = setXmlAttribute(nextTag, 'customHeight', formatXmlBoolean(row.customHeight ?? true))
    }
  } else if (row.customHeight !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'customHeight', formatXmlBoolean(row.customHeight))
  }
  if (row.hidden !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'hidden', formatXmlBoolean(row.hidden))
  }
  if (row.outlineLevel !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'outlineLevel', formatXmlNumber(row.outlineLevel))
  }
  if (row.collapsed !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'collapsed', formatXmlBoolean(row.collapsed))
  }
  if (row.thickTop !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'thickTop', formatXmlBoolean(row.thickTop))
  }
  if (row.thickBottom !== undefined) {
    nextTag = setXmlAttribute(nextTag, 'thickBot', formatXmlBoolean(row.thickBottom))
  }
  return nextTag
}

function buildEmptyRowXml(row: ExportRowMetadata): string {
  let rowTag = `<row r="${escapeXmlAttribute(String(row.rowNumber))}"/>`
  rowTag = applyRowMetadata(rowTag, row)
  return rowTag
}

function updateExistingRowXml(rowXml: string, row: ExportRowMetadata): string {
  const openingTag = worksheetRowOpeningTagPattern.exec(rowXml)?.[0]
  if (!openingTag) {
    return rowXml
  }
  if (openingTag.endsWith('/>')) {
    return applyRowMetadata(openingTag, row)
  }
  const rowBody = rowXml.slice(openingTag.length, -'</row>'.length)
  return rowBody.trim().length === 0 ? buildEmptyRowXml(row) : `${applyRowMetadata(openingTag, row)}${rowBody}</row>`
}

function upsertWorksheetRowMetadata(sheetXml: string, rows: readonly ExportRowMetadata[]): string {
  const rowsByNumber = new Map(rows.map((row) => [row.rowNumber, row]))
  const sortedMissingRows = [...rowsByNumber.values()].toSorted((left, right) => left.rowNumber - right.rowNumber)
  const selfClosingSheetDataMatch = /<sheetData\b([^>]*)\/>/u.exec(sheetXml)
  if (selfClosingSheetDataMatch) {
    const rowXml = sortedMissingRows.map(buildEmptyRowXml).join('')
    return sheetXml.replace(/<sheetData\b([^>]*)\/>/u, (_match, attributes: string) => `<sheetData${attributes}>${rowXml}</sheetData>`)
  }

  const sheetDataMatch = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u.exec(sheetXml)
  if (!sheetDataMatch) {
    return sheetXml
  }

  const sheetDataXml = sheetDataMatch[0]
  const sheetDataOpeningTag = /^<sheetData\b[^>]*>/u.exec(sheetDataXml)?.[0]
  if (!sheetDataOpeningTag || !sheetDataXml.endsWith('</sheetData>')) {
    return sheetXml
  }

  const bodyStart = sheetDataOpeningTag.length
  const bodyEnd = sheetDataXml.length - '</sheetData>'.length
  const sheetDataBody = sheetDataXml.slice(bodyStart, bodyEnd)
  let outputBody = ''
  let lastIndex = 0
  let missingIndex = 0
  for (const match of sheetDataBody.matchAll(worksheetRowElementPattern)) {
    const rowXml = match[0]
    const existingRowNumber = readRowNumber(rowXml)
    outputBody += sheetDataBody.slice(lastIndex, match.index)
    if (existingRowNumber !== null) {
      while (missingIndex < sortedMissingRows.length && sortedMissingRows[missingIndex]!.rowNumber < existingRowNumber) {
        outputBody += buildEmptyRowXml(sortedMissingRows[missingIndex]!)
        missingIndex += 1
      }
      const row = rowsByNumber.get(existingRowNumber)
      outputBody += row ? updateExistingRowXml(rowXml, row) : rowXml
      if (row) {
        while (missingIndex < sortedMissingRows.length && sortedMissingRows[missingIndex]!.rowNumber <= existingRowNumber) {
          missingIndex += 1
        }
      }
    } else {
      outputBody += rowXml
    }
    lastIndex = match.index + rowXml.length
  }
  outputBody += sheetDataBody.slice(lastIndex)
  while (missingIndex < sortedMissingRows.length) {
    outputBody += buildEmptyRowXml(sortedMissingRows[missingIndex]!)
    missingIndex += 1
  }
  const updatedSheetDataXml = `${sheetDataOpeningTag}${outputBody}</sheetData>`
  return `${sheetXml.slice(0, sheetDataMatch.index)}${updatedSheetDataXml}${sheetXml.slice(sheetDataMatch.index + sheetDataXml.length)}`
}

export function hasExportWorksheetDimensions(snapshot: WorkbookSnapshot): boolean {
  return snapshot.sheets.some((sheet) => {
    const metadata = sheet.metadata
    return (
      metadata?.sheetFormatPr !== undefined ||
      normalizeExportColumnMetadata(metadata?.columnMetadata).length > 0 ||
      normalizeExportRowMetadata(metadata?.rows, metadata?.rowMetadata).length > 0
    )
  })
}

export function applyExportWorksheetDimensionsToWorksheetXml(sheetXml: string, metadata: SheetMetadataSnapshot | undefined): string {
  const sheetFormatXml = applySheetFormatPr(sheetXml, metadata?.sheetFormatPr)
  const columnXml = applyColumnMetadata(sheetFormatXml, normalizeExportColumnMetadata(metadata?.columnMetadata))
  const normalizedRows = normalizeExportRowMetadata(metadata?.rows, metadata?.rowMetadata)
  return normalizedRows.length > 0 ? upsertWorksheetRowMetadata(columnXml, normalizedRows) : columnXml
}

export function addExportWorksheetDimensionsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!hasExportWorksheetDimensions(snapshot)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const updatedSheetXml = applyExportWorksheetDimensionsToWorksheetXml(sheetXml, sheet.metadata)
      if (updatedSheetXml === sheetXml) {
        return
      }
      setZipText(zip, sheetPath, updatedSheetXml)
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
