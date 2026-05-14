import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'
import type * as XLSXStyle from 'xlsx-js-style'

import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleBordersSnapshot,
  CellStyleFillSnapshot,
  CellStyleFontSnapshot,
  CellStyleProtectionSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
  WorkbookAxisMetadataSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookSheetFormatPrSnapshot,
  WorkbookSheetStyleArtifactsSnapshot,
  WorkbookStyleArtifactsSnapshot,
  SheetStyleRangeSnapshot,
} from '@bilig/protocol'
import { readImportedWorkbookThemeArtifact } from './xlsx-theme-artifacts.js'
import { workbookDirectorySheetPaths, workbookSheetPath, workbookSheetPathsByName } from './xlsx-workbook-sheet-paths.js'
import { getZipText as getZipEntryText, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ImportedCellStyle = Omit<CellStyleRecord, 'id'>
type ExportCellAlignment = Record<string, boolean | number | string>

interface ImportedSheetDimensions {
  columns?: WorkbookAxisEntrySnapshot[]
  rows?: WorkbookAxisEntrySnapshot[]
  columnMetadata?: WorkbookAxisMetadataSnapshot[]
  rowMetadata?: WorkbookAxisMetadataSnapshot[]
  sheetFormatPr?: WorkbookSheetFormatPrSnapshot
  skippedColumnMetadata?: boolean
}

interface ImportedWorkbookFileStylesOptions {
  styleCandidateAddressesBySheet?: ReadonlyMap<string, ReadonlySet<string>>
}

interface ImportedWorkbookStyleArtifacts {
  workbookArtifacts?: WorkbookStyleArtifactsSnapshot
  sheetArtifactsByName: Map<string, WorkbookSheetStyleArtifactsSnapshot>
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

// XLSX can encode whole-sheet visual defaults as one <col min="1" max="16384"> range.
// Preserve bounded column metadata, but do not expand broad defaults into snapshot state.
const maxExpandedColumnMetadataEntries = 2_048

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function stringValue(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

function numberValue(value: unknown): number | null {
  if (typeof value !== 'string' || value.trim().length === 0) {
    return null
  }
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getFileText(files: unknown, path: string): string | null {
  if (!isRecord(files)) {
    return null
  }
  const file = files[normalizeZipPath(path)]
  if (!isRecord(file)) {
    return null
  }
  const content = file['content']
  if (typeof content === 'string') {
    return content
  }
  if (content instanceof ArrayBuffer) {
    return new TextDecoder().decode(content)
  }
  if (ArrayBuffer.isView(content)) {
    return new TextDecoder().decode(content)
  }
  return null
}

function getPackageText(files: unknown, sourceZip: XlsxZipSource | null, path: string): string | null {
  return sourceZip ? getZipEntryText(readXlsxZipEntries(sourceZip), path) : getFileText(files, path)
}

function workbookRecord(workbook: XLSX.WorkBook): Record<string, unknown> | null {
  const value: unknown = workbook
  return isRecord(value) ? value : null
}

function workbookFiles(workbook: XLSX.WorkBook): unknown {
  return workbookRecord(workbook)?.['files']
}

function workbookStylePath(workbook: XLSX.WorkBook): string | null {
  const directory = workbookRecord(workbook)?.['Directory']
  if (!isRecord(directory)) {
    return null
  }
  if (typeof directory['style'] === 'string') {
    return directory['style']
  }
  const firstStylePath = asArray(directory['styles']).find((entry) => typeof entry === 'string')
  return typeof firstStylePath === 'string' ? firstStylePath : null
}

function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') {
    return null
  }
  const normalized = value.trim().replace(/^#/, '')
  if (/^[0-9a-fA-F]{6}$/.test(normalized)) {
    return `#${normalized.toLowerCase()}`
  }
  if (/^[0-9a-fA-F]{8}$/.test(normalized)) {
    return `#${normalized.slice(2).toLowerCase()}`
  }
  return null
}

function toArgbColor(value: string): string | null {
  const normalized = normalizeRgbColor(value)
  return normalized ? `FF${normalized.slice(1).toUpperCase()}` : null
}

function readColorRecord(value: unknown): string | null {
  return normalizeRgbColor(
    recordChild(value, 'color')?.['rgb'] ?? recordChild(value, 'fgColor')?.['rgb'] ?? recordChild(value, 'bgColor')?.['rgb'],
  )
}

function readFont(font: unknown): CellStyleFontSnapshot | undefined {
  if (!isRecord(font)) {
    return undefined
  }
  const output: CellStyleFontSnapshot = {}
  const name = stringValue(recordChild(font, 'name')?.['val'])
  const size = numberValue(recordChild(font, 'sz')?.['val'])
  const color = readColorRecord(font)
  if (name && name.length > 0) {
    output.family = name
  }
  if (size !== null && size > 0) {
    output.size = size
  }
  if (font['b'] !== undefined) {
    output.bold = true
  }
  if (font['i'] !== undefined) {
    output.italic = true
  }
  if (font['u'] !== undefined) {
    output.underline = true
  }
  if (color) {
    output.color = color
  }
  return Object.keys(output).length > 0 ? output : undefined
}

function readFill(fill: unknown): CellStyleFillSnapshot | undefined {
  const patternFill = recordChild(fill, 'patternFill')
  if (!patternFill || patternFill['patternType'] !== 'solid') {
    return undefined
  }
  const backgroundColor = normalizeRgbColor(recordChild(patternFill, 'fgColor')?.['rgb'] ?? recordChild(patternFill, 'bgColor')?.['rgb'])
  return backgroundColor ? { backgroundColor } : undefined
}

function readBorderKind(value: unknown): { style: CellBorderStyle; weight: CellBorderWeight } | null {
  switch (value) {
    case 'hair':
    case 'thin':
      return { style: 'solid', weight: 'thin' }
    case 'medium':
      return { style: 'solid', weight: 'medium' }
    case 'thick':
      return { style: 'solid', weight: 'thick' }
    case 'dashed':
      return { style: 'dashed', weight: 'thin' }
    case 'mediumDashed':
      return { style: 'dashed', weight: 'medium' }
    case 'dotted':
      return { style: 'dotted', weight: 'thin' }
    default:
      return null
  }
}

function readBorderSide(side: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(side)) {
    return undefined
  }
  const borderKind = readBorderKind(side['style'])
  if (!borderKind) {
    return undefined
  }
  return {
    ...borderKind,
    color: normalizeRgbColor(recordChild(side, 'color')?.['rgb']) ?? '#000000',
  }
}

function readBorders(border: unknown): CellStyleBordersSnapshot | undefined {
  if (!isRecord(border)) {
    return undefined
  }
  const top = readBorderSide(border['top'])
  const right = readBorderSide(border['right'])
  const bottom = readBorderSide(border['bottom'])
  const left = readBorderSide(border['left'])
  const borders: CellStyleBordersSnapshot = {}
  if (top) {
    borders.top = top
  }
  if (right) {
    borders.right = right
  }
  if (bottom) {
    borders.bottom = bottom
  }
  if (left) {
    borders.left = left
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function readHorizontalAlignment(value: unknown): CellHorizontalAlignment | undefined {
  switch (value) {
    case 'general':
    case 'left':
    case 'center':
    case 'right':
    case 'fill':
    case 'justify':
    case 'centerContinuous':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function readVerticalAlignment(value: unknown): CellVerticalAlignment | undefined {
  switch (value) {
    case 'top':
      return 'top'
    case 'center':
    case 'middle':
      return 'middle'
    case 'bottom':
      return 'bottom'
    case 'justify':
      return 'justify'
    case 'distributed':
      return 'distributed'
    default:
      return undefined
  }
}

function readAlignment(alignment: unknown): CellStyleAlignmentSnapshot | undefined {
  if (!isRecord(alignment)) {
    return undefined
  }
  const horizontal = readHorizontalAlignment(alignment['horizontal'])
  const vertical = readVerticalAlignment(alignment['vertical'])
  const indent = numberValue(alignment['indent'])
  const readingOrder = numberValue(alignment['readingOrder'])
  const textRotation = numberValue(alignment['textRotation'])
  const output: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(readBooleanAttribute(alignment['wrapText']) === true ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
    ...(readBooleanAttribute(alignment['shrinkToFit']) === true ? { shrinkToFit: true } : {}),
    ...(readingOrder !== null ? { readingOrder } : {}),
    ...(textRotation !== null ? { textRotation } : {}),
    ...(readBooleanAttribute(alignment['justifyLastLine']) === true ? { justifyLastLine: true } : {}),
  }
  return Object.keys(output).length > 0 ? output : undefined
}

function readBooleanAttribute(value: unknown): boolean | undefined {
  if (value === true || value === '1' || (typeof value === 'string' && value.toLowerCase() === 'true')) {
    return true
  }
  if (value === false || value === '0' || (typeof value === 'string' && value.toLowerCase() === 'false')) {
    return false
  }
  return undefined
}

function isStyleComponentApplied(flag: unknown, componentId: number | null): boolean {
  const explicit = readBooleanAttribute(flag)
  return explicit ?? (componentId !== null && componentId > 0)
}

function readProtection(protection: unknown): CellStyleProtectionSnapshot | undefined {
  if (!isRecord(protection)) {
    return undefined
  }
  const locked = readBooleanAttribute(protection['locked'])
  const hidden = readBooleanAttribute(protection['hidden'])
  return {
    ...(locked !== undefined ? { locked } : {}),
    ...(hidden !== undefined ? { hidden } : {}),
  }
}

function parseWorkbookStyles(stylesXml: string): Map<number, ImportedCellStyle> {
  const relevantStylesXml = ['fonts', 'fills', 'borders', 'cellXfs']
    .flatMap((elementName) => {
      const elementXml = extractStyleXmlElement(stylesXml, elementName)
      return elementXml ? [elementXml] : []
    })
    .join('')
  if (relevantStylesXml.length === 0) {
    return new Map()
  }
  const parsed: unknown = xmlParser.parse(`<styleSheet>${relevantStylesXml}</styleSheet>`)
  const styleSheet = recordChild(parsed, 'styleSheet')
  if (!styleSheet) {
    return new Map()
  }

  const fonts = asArray(recordChild(styleSheet, 'fonts')?.['font']).map(readFont)
  const fills = asArray(recordChild(styleSheet, 'fills')?.['fill']).map(readFill)
  const borders = asArray(recordChild(styleSheet, 'borders')?.['border']).map(readBorders)
  const cellXfs = asArray(recordChild(styleSheet, 'cellXfs')?.['xf'])
  const styles = new Map<number, ImportedCellStyle>()

  cellXfs.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const fontId = numberValue(entry['fontId'])
    const fillId = numberValue(entry['fillId'])
    const borderId = numberValue(entry['borderId'])
    const font = fontId !== null ? fonts[fontId] : undefined
    const fill = fillId !== null ? fills[fillId] : undefined
    const bordersValue = borderId !== null ? borders[borderId] : undefined
    const alignment = readAlignment(entry['alignment'])
    const protection = entry['applyProtection'] === '1' ? (readProtection(entry['protection']) ?? {}) : readProtection(entry['protection'])
    const style: ImportedCellStyle = {
      ...(isStyleComponentApplied(entry['applyFill'], fillId) && fill ? { fill } : {}),
      ...(isStyleComponentApplied(entry['applyFont'], fontId) && font ? { font } : {}),
      ...(entry['applyAlignment'] === '1' && alignment ? { alignment } : {}),
      ...(isStyleComponentApplied(entry['applyBorder'], borderId) && bordersValue ? { borders: bordersValue } : {}),
      ...(protection !== undefined ? { protection } : {}),
    }
    if (Object.keys(style).length > 0) {
      styles.set(index, style)
    }
  })

  return styles
}

function extractStyleXmlElement(stylesXml: string, elementName: string): string | null {
  const qualifiedName = `(?:[A-Za-z_][\\w.-]*:)?${elementName}`
  const expanded = new RegExp(`<${qualifiedName}\\b[^>]*>[\\s\\S]*?<\\/${qualifiedName}>`, 'u').exec(stylesXml)
  if (expanded) {
    return expanded[0]
  }
  const selfClosing = new RegExp(`<${qualifiedName}\\b[^>]*\\/>`, 'u').exec(stylesXml)
  return selfClosing?.[0] ?? null
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
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

function readXmlPositiveIntegerAttribute(tag: string, name: string): number | null {
  const value = readXmlNumberAttribute(tag, name)
  return value !== null && Number.isSafeInteger(value) && value > 0 ? value : null
}

function readXmlNonNegativeIntegerAttribute(tag: string, name: string): number | null {
  const value = readXmlNumberAttribute(tag, name)
  return value !== null && Number.isSafeInteger(value) && value >= 0 ? value : null
}

function readXmlOptionalBooleanAttribute(tag: string, name: string): boolean | null {
  const raw = readXmlAttribute(tag, name)
  if (raw === null) {
    return null
  }
  return raw === '1' || raw.toLowerCase() === 'true'
}

function parseSheetStyleIndexes(sheetXml: string, candidateAddresses?: ReadonlySet<string>): Map<string, number> {
  const output = new Map<string, number>()
  if (candidateAddresses?.size === 0) {
    return output
  }
  let remainingCandidateCount = candidateAddresses?.size ?? null

  for (const match of sheetXml.matchAll(/<c\b[^>]*>/gu)) {
    const cellTag = match[0]
    const address = readXmlAttribute(cellTag, 'r')
    if (!address) {
      continue
    }
    if (candidateAddresses) {
      if (!candidateAddresses.has(address)) {
        continue
      }
      remainingCandidateCount = remainingCandidateCount === null ? null : remainingCandidateCount - 1
    }
    const styleIndexValue = readXmlAttribute(cellTag, 's')
    if (styleIndexValue === null || styleIndexValue.trim().length === 0) {
      if (remainingCandidateCount === 0) {
        break
      }
      continue
    }
    const styleIndex = Number(styleIndexValue)
    if (Number.isSafeInteger(styleIndex)) {
      output.set(address, styleIndex)
    }
    if (remainingCandidateCount === 0) {
      break
    }
  }

  return output
}

function parseColumnStyleRanges(sheetXml: string): Array<{ start: number; end: number }> {
  return [...sheetXml.matchAll(/<col\b[^>]*\/?>/gu)].flatMap((match) => {
    const columnTag = match[0]
    if (readXmlNonNegativeIntegerAttribute(columnTag, 'style') === null) {
      return []
    }
    const min = readXmlPositiveIntegerAttribute(columnTag, 'min')
    const max = readXmlPositiveIntegerAttribute(columnTag, 'max') ?? min
    if (min === null || max === null || max < min) {
      return []
    }
    return [{ start: min - 1, end: max - 1 }]
  })
}

function columnHasStyle(columnStyleRanges: readonly { start: number; end: number }[], columnIndex: number): boolean {
  return columnStyleRanges.some((range) => columnIndex >= range.start && columnIndex <= range.end)
}

function isBlankCellXml(cellXml: string, openingTag: string): boolean {
  if (openingTag.endsWith('/>')) {
    return true
  }
  return !/<(?:v|f|is)\b/u.test(cellXml.slice(openingTag.length, -'</c>'.length))
}

function parseSheetCellStyleIndexArtifacts(sheetXml: string): WorkbookSheetStyleArtifactsSnapshot | undefined {
  const cellStyleIndexes: WorkbookSheetStyleArtifactsSnapshot['cellStyleIndexes'] = []
  const blankCellAddresses: string[] = []
  const columnStyleRanges = parseColumnStyleRanges(sheetXml)
  for (const rowMatch of sheetXml.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gu)) {
    const rowXml = rowMatch[0]
    const rowTag = /^<row\b[^>]*(?:\/>|>)/u.exec(rowXml)?.[0]
    if (!rowTag) {
      continue
    }
    const rowHasStyle = readXmlNonNegativeIntegerAttribute(rowTag, 's') !== null
    for (const cellMatch of rowXml.matchAll(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gu)) {
      const cellXml = cellMatch[0]
      const cellTag = /^<c\b[^>]*(?:\/>|>)/u.exec(cellXml)?.[0]
      if (!cellTag) {
        continue
      }
      const address = readXmlAttribute(cellTag, 'r')
      if (!address) {
        continue
      }
      const styleIndex = readXmlNonNegativeIntegerAttribute(cellTag, 's')
      if (styleIndex !== null) {
        cellStyleIndexes.push({ address, styleIndex })
        continue
      }
      if (!isBlankCellXml(cellXml, cellTag)) {
        continue
      }
      const columnIndex = XLSX.utils.decode_cell(address).c
      if (rowHasStyle || columnHasStyle(columnStyleRanges, columnIndex)) {
        blankCellAddresses.push(address)
      }
    }
  }
  return cellStyleIndexes.length > 0 || blankCellAddresses.length > 0
    ? {
        cellStyleIndexes,
        ...(blankCellAddresses.length > 0 ? { blankCellAddresses } : {}),
      }
    : undefined
}

function parseSheetFormatPr(sheetXml: string): WorkbookSheetFormatPrSnapshot | undefined {
  const match = /<sheetFormatPr\b[^>]*\/?>/u.exec(sheetXml)
  if (!match) {
    return undefined
  }
  const tag = match[0]
  const sheetFormatPr: WorkbookSheetFormatPrSnapshot = {
    ...(readXmlNumberAttribute(tag, 'baseColWidth') !== null ? { baseColWidth: readXmlNumberAttribute(tag, 'baseColWidth') } : {}),
    ...(readXmlNumberAttribute(tag, 'defaultColWidth') !== null ? { defaultColWidth: readXmlNumberAttribute(tag, 'defaultColWidth') } : {}),
    ...(readXmlNumberAttribute(tag, 'defaultRowHeight') !== null
      ? { defaultRowHeight: readXmlNumberAttribute(tag, 'defaultRowHeight') }
      : {}),
    ...(readXmlOptionalBooleanAttribute(tag, 'customHeight') !== null
      ? { customHeight: readXmlOptionalBooleanAttribute(tag, 'customHeight') }
      : {}),
    ...(readXmlNonNegativeIntegerAttribute(tag, 'outlineLevelRow') !== null
      ? { outlineLevelRow: readXmlNonNegativeIntegerAttribute(tag, 'outlineLevelRow') }
      : {}),
    ...(readXmlNonNegativeIntegerAttribute(tag, 'outlineLevelCol') !== null
      ? { outlineLevelCol: readXmlNonNegativeIntegerAttribute(tag, 'outlineLevelCol') }
      : {}),
    ...(readXmlOptionalBooleanAttribute(tag, 'thickTop') !== null ? { thickTop: readXmlOptionalBooleanAttribute(tag, 'thickTop') } : {}),
    ...(readXmlOptionalBooleanAttribute(tag, 'thickBottom') !== null
      ? { thickBottom: readXmlOptionalBooleanAttribute(tag, 'thickBottom') }
      : {}),
  }
  return Object.keys(sheetFormatPr).length > 0 ? sheetFormatPr : undefined
}

function parseSheetColumnEntries(sheetXml: string): {
  entries?: WorkbookAxisEntrySnapshot[]
  metadata?: WorkbookAxisMetadataSnapshot[]
  skipped: boolean
} {
  const entries: WorkbookAxisEntrySnapshot[] = []
  const metadata: WorkbookAxisMetadataSnapshot[] = []
  let skipped = false
  for (const match of sheetXml.matchAll(/<col\b[^>]*\/?>/gu)) {
    const columnTag = match[0]
    const min = readXmlPositiveIntegerAttribute(columnTag, 'min')
    const max = readXmlPositiveIntegerAttribute(columnTag, 'max') ?? min
    if (min === null || max === null || max < min) {
      continue
    }
    const width = readXmlNumberAttribute(columnTag, 'width')
    const widthSize = width !== null && width > 0 ? Math.round(width * 6) : null
    const size = widthSize !== null && widthSize > 0 ? widthSize : null
    const startColumn = min - 1
    const endColumn = max - 1
    const columnCount = endColumn - startColumn + 1
    if (columnCount <= 0) {
      continue
    }
    const customWidth = readXmlOptionalBooleanAttribute(columnTag, 'customWidth')
    const styleIndex = readXmlNonNegativeIntegerAttribute(columnTag, 'style')
    const customFormat = readXmlOptionalBooleanAttribute(columnTag, 'customFormat')
    const bestFit = readXmlOptionalBooleanAttribute(columnTag, 'bestFit')
    const hidden = readXmlOptionalBooleanAttribute(columnTag, 'hidden')
    const outlineLevel = readXmlNonNegativeIntegerAttribute(columnTag, 'outlineLevel')
    const collapsed = readXmlOptionalBooleanAttribute(columnTag, 'collapsed')
    if (
      size === null &&
      customWidth === null &&
      styleIndex === null &&
      customFormat === null &&
      bestFit === null &&
      hidden === null &&
      outlineLevel === null &&
      collapsed === null
    ) {
      continue
    }
    metadata.push({
      start: startColumn,
      count: columnCount,
      ...(size !== null && size > 0 ? { size } : {}),
      ...(width !== null && width > 0 ? { xlsxWidth: width } : {}),
      ...(styleIndex !== null ? { styleIndex } : {}),
      ...(customFormat !== null ? { customFormat } : {}),
      ...(customWidth !== null ? { customWidth } : {}),
      ...(bestFit !== null ? { bestFit } : {}),
      ...(hidden !== null ? { hidden } : {}),
      ...(outlineLevel !== null ? { outlineLevel } : {}),
      ...(collapsed !== null ? { collapsed } : {}),
    })
    if (size === null && hidden !== true) {
      continue
    }
    if (entries.length + columnCount > maxExpandedColumnMetadataEntries) {
      skipped = true
      continue
    }
    for (let column = startColumn; column <= endColumn; column += 1) {
      entries.push({
        id: `col:${String(column)}`,
        index: column,
        ...(size !== null ? { size } : {}),
        ...(hidden === true ? { hidden: true } : {}),
      })
    }
  }
  return {
    ...(entries.length > 0 ? { entries } : {}),
    ...(metadata.length > 0 ? { metadata } : {}),
    skipped,
  }
}

function rowMetadataValuesMatch(left: WorkbookAxisMetadataSnapshot, right: WorkbookAxisMetadataSnapshot): boolean {
  return (
    left.size === right.size &&
    left.hidden === right.hidden &&
    left.xlsxHeight === right.xlsxHeight &&
    left.styleIndex === right.styleIndex &&
    left.customFormat === right.customFormat &&
    left.customHeight === right.customHeight &&
    left.outlineLevel === right.outlineLevel &&
    left.collapsed === right.collapsed &&
    left.thickTop === right.thickTop &&
    left.thickBottom === right.thickBottom
  )
}

function coalesceRowMetadata(records: readonly WorkbookAxisMetadataSnapshot[]): WorkbookAxisMetadataSnapshot[] {
  const coalesced: WorkbookAxisMetadataSnapshot[] = []
  for (const record of records) {
    const previous = coalesced[coalesced.length - 1]
    if (previous && previous.start + previous.count === record.start && rowMetadataValuesMatch(previous, record)) {
      coalesced[coalesced.length - 1] = { ...previous, count: previous.count + record.count }
      continue
    }
    coalesced.push({ ...record })
  }
  return coalesced
}

function parseSheetRowEntries(sheetXml: string): { entries?: WorkbookAxisEntrySnapshot[]; metadata?: WorkbookAxisMetadataSnapshot[] } {
  const entries: WorkbookAxisEntrySnapshot[] = []
  const metadata: WorkbookAxisMetadataSnapshot[] = []
  for (const match of sheetXml.matchAll(/<row\b[^>]*>/gu)) {
    const rowTag = match[0]
    const rowNumber = readXmlPositiveIntegerAttribute(rowTag, 'r')
    const height = readXmlNumberAttribute(rowTag, 'ht')
    const hidden = readXmlOptionalBooleanAttribute(rowTag, 'hidden')
    if (rowNumber === null) {
      continue
    }
    const index = rowNumber - 1
    const size = height !== null && height > 0 ? Math.round(height) : null
    const styleIndex = readXmlNonNegativeIntegerAttribute(rowTag, 's')
    const customFormat = readXmlOptionalBooleanAttribute(rowTag, 'customFormat')
    const customHeight = readXmlOptionalBooleanAttribute(rowTag, 'customHeight')
    const outlineLevel = readXmlNonNegativeIntegerAttribute(rowTag, 'outlineLevel')
    const collapsed = readXmlOptionalBooleanAttribute(rowTag, 'collapsed')
    const thickTop = readXmlOptionalBooleanAttribute(rowTag, 'thickTop')
    const thickBottom = readXmlOptionalBooleanAttribute(rowTag, 'thickBot')
    if (
      (size === null || size <= 0) &&
      hidden !== true &&
      styleIndex === null &&
      customFormat === null &&
      customHeight === null &&
      outlineLevel === null &&
      collapsed === null &&
      thickTop === null &&
      thickBottom === null
    ) {
      continue
    }
    if (size !== null || hidden === true) {
      entries.push({
        id: `row:${String(index)}`,
        index,
        ...(size !== null && size > 0 ? { size } : {}),
        ...(hidden === true ? { hidden: true } : {}),
      })
    }
    metadata.push({
      start: index,
      count: 1,
      ...(size !== null && size > 0 ? { size } : {}),
      ...(height !== null && height > 0 ? { xlsxHeight: height } : {}),
      ...(hidden !== null ? { hidden } : {}),
      ...(styleIndex !== null ? { styleIndex } : {}),
      ...(customFormat !== null ? { customFormat } : {}),
      ...(customHeight !== null ? { customHeight } : {}),
      ...(outlineLevel !== null ? { outlineLevel } : {}),
      ...(collapsed !== null ? { collapsed } : {}),
      ...(thickTop !== null ? { thickTop } : {}),
      ...(thickBottom !== null ? { thickBottom } : {}),
    })
  }
  const coalescedMetadata = coalesceRowMetadata(metadata)
  return {
    ...(entries.length > 0 ? { entries } : {}),
    ...(coalescedMetadata.length > 0 ? { metadata: coalescedMetadata } : {}),
  }
}

export function readImportedWorkbookSheetDimensions(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
): Map<string, ImportedSheetDimensions> {
  const files = workbookFiles(workbook)
  const sheetPathsByName = workbookSheetPathsByName(workbook)
  const fallbackSheetPaths = workbookDirectorySheetPaths(workbook)
  const output = new Map<string, ImportedSheetDimensions>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    const sheetXml = sheetPath ? getFileText(files, sheetPath) : null
    if (!sheetXml) {
      return
    }
    const sheetFormatPr = parseSheetFormatPr(sheetXml)
    const parsedColumns = parseSheetColumnEntries(sheetXml)
    const parsedRows = parseSheetRowEntries(sheetXml)
    const dimensions: ImportedSheetDimensions = {
      ...(parsedColumns.entries ? { columns: parsedColumns.entries } : {}),
      ...(parsedRows.entries ? { rows: parsedRows.entries } : {}),
      ...(parsedColumns.metadata ? { columnMetadata: parsedColumns.metadata } : {}),
      ...(parsedRows.metadata ? { rowMetadata: parsedRows.metadata } : {}),
      ...(sheetFormatPr ? { sheetFormatPr } : {}),
      ...(parsedColumns.skipped ? { skippedColumnMetadata: true } : {}),
    }
    if (
      dimensions.columns ||
      dimensions.rows ||
      dimensions.columnMetadata ||
      dimensions.rowMetadata ||
      dimensions.sheetFormatPr ||
      dimensions.skippedColumnMetadata
    ) {
      output.set(sheetName, dimensions)
    }
  })
  return output
}

export function readImportedWorkbookFileStyles(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  options: ImportedWorkbookFileStylesOptions = {},
  source?: XlsxZipSource,
): Map<string, Map<string, ImportedCellStyle>> {
  const files = workbookFiles(workbook)
  const sourceZip = source ? readXlsxZipEntries(source) : null
  const stylePath = workbookStylePath(workbook)
  const styleXml = stylePath ? getPackageText(files, sourceZip, stylePath) : null
  if (!styleXml) {
    return new Map()
  }
  const stylesByIndex = parseWorkbookStyles(styleXml)
  if (stylesByIndex.size === 0) {
    return new Map()
  }

  const sheetPathsByName = workbookSheetPathsByName(workbook)
  const fallbackSheetPaths = workbookDirectorySheetPaths(workbook)
  const output = new Map<string, Map<string, ImportedCellStyle>>()
  sheetNames.forEach((sheetName, index) => {
    const candidateAddresses = options.styleCandidateAddressesBySheet?.get(sheetName)
    if (candidateAddresses?.size === 0) {
      return
    }
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    const sheetXml = sheetPath ? getPackageText(files, sourceZip, sheetPath) : null
    if (!sheetXml) {
      return
    }
    const styleIndexes = parseSheetStyleIndexes(sheetXml, candidateAddresses)
    const cellStyles = new Map<string, ImportedCellStyle>()
    for (const [address, styleIndex] of styleIndexes) {
      const style = stylesByIndex.get(styleIndex)
      if (style) {
        cellStyles.set(address, style)
      }
    }
    if (cellStyles.size > 0) {
      output.set(sheetName, cellStyles)
    }
  })

  return output
}

export function readImportedWorkbookStyleArtifacts(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  source?: XlsxZipSource,
): ImportedWorkbookStyleArtifacts {
  const zip = source ? readXlsxZipEntries(source) : null
  const files = workbookFiles(workbook)
  const stylePath = workbookStylePath(workbook)
  const readPartText = (path: string | null | undefined): string | null => {
    if (!path) {
      return null
    }
    return zip ? getZipEntryText(zip, path) : getFileText(files, path)
  }
  const stylesXml = readPartText(stylePath)
  const theme = readImportedWorkbookThemeArtifact(zip ?? undefined)
  const sheetPathsByName = workbookSheetPathsByName(workbook)
  const fallbackSheetPaths = workbookDirectorySheetPaths(workbook)
  const sheetArtifactsByName = new Map<string, WorkbookSheetStyleArtifactsSnapshot>()

  sheetNames.forEach((sheetName, index) => {
    const sheetPath = workbookSheetPath(sheetPathsByName, fallbackSheetPaths, sheetName, index)
    const sheetXml = readPartText(sheetPath)
    if (!sheetXml) {
      return
    }
    const sheetArtifacts = parseSheetCellStyleIndexArtifacts(sheetXml)
    if (sheetArtifacts) {
      sheetArtifactsByName.set(sheetName, sheetArtifacts)
    }
  })

  return {
    ...(stylesXml ? { workbookArtifacts: { stylesXml, ...(theme ? { theme } : {}) } } : {}),
    sheetArtifactsByName,
  }
}

function exportColor(value: string): XLSXStyle.CellStyleColor | undefined {
  const rgb = toArgbColor(value)
  return rgb ? { rgb } : undefined
}

function exportAlignment(alignment: CellStyleAlignmentSnapshot): ExportCellAlignment {
  return {
    ...(alignment.horizontal ? { horizontal: alignment.horizontal } : {}),
    ...(alignment.vertical ? { vertical: alignment.vertical === 'middle' ? 'center' : alignment.vertical } : {}),
    ...(alignment.wrap === true ? { wrapText: true } : {}),
    ...(alignment.indent !== undefined && alignment.indent >= 0 ? { indent: alignment.indent } : {}),
    ...(alignment.shrinkToFit === true ? { shrinkToFit: true } : {}),
    ...(alignment.readingOrder !== undefined ? { readingOrder: alignment.readingOrder } : {}),
    ...(alignment.textRotation !== undefined ? { textRotation: alignment.textRotation } : {}),
    ...(alignment.justifyLastLine === true ? { justifyLastLine: true } : {}),
  }
}

function exportBorderSide(side: CellBorderSideSnapshot): { color: XLSXStyle.CellStyleColor; style?: XLSXStyle.BorderType } {
  const color = exportColor(side.color) ?? { rgb: 'FF000000' }
  switch (side.style) {
    case 'solid':
      return { color, style: side.weight }
    case 'dashed':
      return { color, style: side.weight === 'thin' ? 'dashed' : 'mediumDashed' }
    case 'dotted':
      return { color, style: 'dotted' }
    case 'double':
      return { color, style: 'thick' }
  }
}

function exportStyle(style: CellStyleRecord): XLSXStyle.CellStyle {
  const output: XLSXStyle.CellStyle = {}
  const fillColor = style.fill?.backgroundColor ? exportColor(style.fill.backgroundColor) : undefined
  if (fillColor) {
    output.fill = { patternType: 'solid', fgColor: fillColor }
  }

  if (style.font) {
    const font: NonNullable<XLSXStyle.CellStyle['font']> = {}
    const fontColor = style.font.color ? exportColor(style.font.color) : undefined
    if (style.font.family) {
      font.name = style.font.family
    }
    if (style.font.size) {
      font.sz = style.font.size
    }
    if (style.font.bold === true) {
      font.bold = true
    }
    if (style.font.italic === true) {
      font.italic = true
    }
    if (style.font.underline === true) {
      font.underline = true
    }
    if (fontColor) {
      font.color = fontColor
    }
    if (Object.keys(font).length > 0) {
      output.font = font
    }
  }

  if (style.alignment) {
    const alignment = exportAlignment(style.alignment)
    if (Object.keys(alignment).length > 0) {
      output.alignment = alignment as NonNullable<XLSXStyle.CellStyle['alignment']>
    }
  }

  if (style.borders) {
    const border: NonNullable<XLSXStyle.CellStyle['border']> = {}
    if (style.borders.top) {
      border.top = exportBorderSide(style.borders.top)
    }
    if (style.borders.right) {
      border.right = exportBorderSide(style.borders.right)
    }
    if (style.borders.bottom) {
      border.bottom = exportBorderSide(style.borders.bottom)
    }
    if (style.borders.left) {
      border.left = exportBorderSide(style.borders.left)
    }
    if (Object.keys(border).length > 0) {
      output.border = border
    }
  }

  return output
}

function isStyledWorksheetCell(value: unknown): value is XLSXStyle.CellObject {
  return isRecord(value) && typeof value['t'] === 'string'
}

export function addExportStylesToWorksheet(
  worksheet: XLSXStyle.WorkSheet,
  styleRanges: readonly SheetStyleRangeSnapshot[] | undefined,
  styles: readonly CellStyleRecord[] | undefined,
): void {
  if (!styleRanges || styleRanges.length === 0 || !styles || styles.length === 0) {
    return
  }

  const stylesById = new Map(styles.map((style) => [style.id, style]))
  for (const styleRange of styleRanges) {
    const style = stylesById.get(styleRange.styleId)
    if (!style) {
      continue
    }
    const exportStyleValue = exportStyle(style)
    const range = XLSX.utils.decode_range(`${styleRange.range.startAddress}:${styleRange.range.endAddress}`)
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: column })
        const existingCell = worksheet[address]
        const cell = isStyledWorksheetCell(existingCell) ? existingCell : ({ t: 'z' } satisfies XLSXStyle.CellObject)
        cell.s = exportStyleValue
        worksheet[address] = cell
      }
    }
  }
}
