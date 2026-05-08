import * as XLSX from 'xlsx'
import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from 'fflate'

import { parseCsv, parseCsvCellInput, resolveCsvParseOptions, type CsvParseOptions } from '@bilig/core'
import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleBordersSnapshot,
  CellStyleFontSnapshot,
  CellStyleProtectionSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookMetadataSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readImportedArrayFormulaSpills } from './xlsx-array-formulas.js'
import { readImportedWorkbookCalculationSettings } from './xlsx-calculation-settings.js'
import { readImportedWorkbookCharts } from './xlsx-charts.js'
import { readImportedSheetComments } from './xlsx-comments.js'
import { readImportedWorkbookConditionalFormats } from './xlsx-conditional-formats.js'
import { readImportedDefinedNames } from './xlsx-defined-names.js'
import { readImportedWorkbookFilters } from './xlsx-filters.js'
import { readImportedWorkbookFreezePanes } from './xlsx-freeze-panes.js'
import { buildMergeEntries } from './xlsx-merge-entries.js'
import { readImportedWorkbookPivots } from './xlsx-pivots.js'
import { readImportedWorkbookProtectedRanges } from './xlsx-protected-ranges.js'
import { readImportedWorkbookSheetProtections } from './xlsx-sheet-protection.js'
import { readImportedWorkbookSorts } from './xlsx-sorts.js'
import { mergeStyleRuns, styleRunsToRanges, type HorizontalStyleRun, type RectangularStyleRun } from './xlsx-style-runs.js'
import { readImportedWorkbookFileStyles, readImportedWorkbookSheetDimensions } from './xlsx-styles.js'
import { readImportedWorkbookSheetTabColors } from './xlsx-tab-colors.js'
import { readImportedWorkbookTables } from './xlsx-tables.js'
import { readImportedWorkbookDataValidations } from './xlsx-validations.js'
import { readImportedWorkbookProperties } from './xlsx-workbook-properties.js'
import {
  createSheetPreview,
  normalizeCsvSheetName,
  normalizeWorkbookName,
  readImportedAlignmentBoolean,
  readImportedAlignmentNumber,
  toDisplayText,
  toLiteralInput,
  type ImportedWorkbookSheetPreview,
} from './workbook-import-helpers.js'
import { readImportedExternalLinkCaches, translateImportedFormulaExternalReferences } from './xlsx-external-references.js'
import { translateImportedFormulaStructuredReferences } from './xlsx-formula-translation.js'
import { createPreservedVbaProjectPayload, type PreservedVbaProjectCodeNames } from './xlsx-macros.js'
import { worksheetCellAt, worksheetCellEntries, worksheetCellEntriesAtAddresses, worksheetCellRecords } from './xlsx-worksheet-cells.js'

export { exportXlsx } from './xlsx-export.js'
export type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'

export const XLSX_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
export const XLSB_CONTENT_TYPE = 'application/vnd.ms-excel.sheet.binary.macroenabled.12'
export const CSV_CONTENT_TYPE = 'text/csv'
export const EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES = [XLSX_CONTENT_TYPE, XLSB_CONTENT_TYPE] as const
export type ExcelWorkbookImportContentType = (typeof EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES)[number]
export const WORKBOOK_IMPORT_CONTENT_TYPES = [...EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES, CSV_CONTENT_TYPE] as const
export type WorkbookImportContentType = (typeof WORKBOOK_IMPORT_CONTENT_TYPES)[number]

export function normalizeWorkbookImportContentType(contentType: string): WorkbookImportContentType | null {
  const mediaType = contentType.split(';', 1)[0]?.trim().toLowerCase() ?? ''
  if (mediaType === XLSX_CONTENT_TYPE || mediaType === XLSB_CONTENT_TYPE || mediaType === CSV_CONTENT_TYPE) {
    return mediaType
  }
  return null
}

const largeWorkbookStyleCandidateThreshold = 100_000
const xlsxWorksheetXmlPathPattern = /^xl\/worksheets\/[^/]+\.xml$/u
const legacyExcelErrorTextByCode = new Map<number, string>([
  [0, '#NULL!'],
  [7, '#DIV/0!'],
  [15, '#VALUE!'],
  [23, '#REF!'],
  [29, '#NAME?'],
  [36, '#NUM!'],
  [42, '#N/A'],
  [43, '#GETTING_DATA'],
])

export interface ImportedWorkbook {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
}

export interface ImportedWorkbookPreview {
  fileName: string
  contentType: WorkbookImportContentType
  fileSizeBytes: number
  workbookName: string
  sheetCount: number
  sheets: readonly ImportedWorkbookSheetPreview[]
  warnings: readonly string[]
}

export type CsvImportOptions = CsvParseOptions

export interface WorkbookImportFileOptions {
  csv?: CsvImportOptions
}

export class InvalidXlsxZipContainerError extends Error {
  constructor() {
    super('Invalid or corrupt XLSX zip container')
    this.name = 'InvalidXlsxZipContainerError'
  }
}

interface SheetColumnInfo {
  index: number
  size: number
}

interface SheetRowInfo {
  index: number
  size: number | null
  hidden: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isStyleOnlyBlankCellTag(tag: string): boolean {
  const attributes = [...tag.matchAll(/\s([A-Za-z_:][\w:.-]*)=(?:"[^"]*"|'[^']*')/gu)].map((match) => match[1])
  const attributeNames = new Set(attributes)
  return attributeNames.size === 2 && attributeNames.has('r') && attributeNames.has('s')
}

function stripStyleOnlyBlankCells(sheetXml: string): string {
  return sheetXml.replace(/<c\b[^>]*\/>/gu, (tag) => (isStyleOnlyBlankCellTag(tag) ? '' : tag))
}

function stripStyleOnlyBlankCellsForSheetJs(data: Uint8Array, zip: Unzipped): Uint8Array {
  let changed = false
  for (const path of Object.keys(zip)) {
    if (!xlsxWorksheetXmlPathPattern.test(path)) {
      continue
    }
    const worksheetBytes = zip[path]
    if (!worksheetBytes) {
      continue
    }
    const worksheetXml = strFromU8(worksheetBytes)
    const strippedWorksheetXml = stripStyleOnlyBlankCells(worksheetXml)
    if (strippedWorksheetXml === worksheetXml) {
      continue
    }
    zip[path] = strToU8(strippedWorksheetXml)
    changed = true
  }
  return changed ? zipSync(zip) : data
}

function createWorkbookPreview(input: {
  contentType: WorkbookImportContentType
  fileName: string
  fileSizeBytes: number
  workbookName: string
  sheets: readonly ImportedWorkbookSheetPreview[]
  warnings: readonly string[]
}): ImportedWorkbookPreview {
  return {
    fileName: input.fileName,
    contentType: input.contentType,
    fileSizeBytes: input.fileSizeBytes,
    workbookName: input.workbookName,
    sheetCount: input.sheets.length,
    sheets: input.sheets,
    warnings: [...input.warnings],
  }
}

function toPixelSize(value: number | undefined, unit: 'pt' | 'ch'): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return null
  }
  if (unit === 'pt') {
    return Math.round((value * 96) / 72)
  }
  return Math.round(value * 8 + 5)
}

function toPositivePixelSize(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? Math.round(value) : null
}

function buildColumnEntries(columns: unknown[] | undefined): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(columns) || columns.length === 0) {
    return undefined
  }
  const entries: SheetColumnInfo[] = []
  columns.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const size =
      typeof entry['wpx'] === 'number'
        ? toPositivePixelSize(entry['wpx'])
        : typeof entry['wch'] === 'number'
          ? toPixelSize(entry['wch'], 'ch')
          : null
    if (size === null) {
      return
    }
    entries.push({ index, size })
  })
  if (entries.length === 0) {
    return undefined
  }
  return entries.map(({ index, size }) => ({
    id: `col:${index}`,
    index,
    size,
  }))
}

function buildRowEntries(rows: unknown[] | undefined): WorkbookAxisEntrySnapshot[] | undefined {
  if (!Array.isArray(rows) || rows.length === 0) {
    return undefined
  }
  const entries: SheetRowInfo[] = []
  rows.forEach((entry, index) => {
    if (!isRecord(entry)) {
      return
    }
    const size =
      typeof entry['hpx'] === 'number'
        ? toPositivePixelSize(entry['hpx'])
        : typeof entry['hpt'] === 'number'
          ? toPixelSize(entry['hpt'], 'pt')
          : null
    const hidden = entry['hidden'] === true
    if (size === null && !hidden) {
      return
    }
    entries.push({ index, size, hidden })
  })
  if (entries.length === 0) {
    return undefined
  }
  return entries.map(({ index, size, hidden }) => {
    const snapshot: WorkbookAxisEntrySnapshot = {
      id: `row:${index}`,
      index,
    }
    if (size !== null) {
      snapshot.size = size
    }
    if (hidden) {
      snapshot.hidden = true
    }
    return snapshot
  })
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

function readRgbColor(value: unknown): string | null {
  if (!isRecord(value)) {
    return null
  }
  return normalizeRgbColor(value['rgb'])
}

function readFiniteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function readImportedNumberFormat(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  if (trimmed.length === 0 || trimmed === 'General') {
    return undefined
  }
  return trimmed
}

function hasImportableXlsxCellPayload(cell: Record<string, unknown>): boolean {
  const formula = cell['f']
  if (typeof formula === 'string' && formula.trim().length > 0) {
    return true
  }
  return toLiteralInput(cell['v']) !== undefined || readImportedNumberFormat(cell['z']) !== undefined
}

function collectStyleCandidateAddresses(
  workbook: XLSX.WorkBook,
  sheetNames: readonly string[],
  maxCandidateCount: number,
): {
  addressesBySheet: Map<string, Set<string>>
  count: number
} {
  const addressesBySheet = new Map<string, Set<string>>()
  let count = 0
  for (const sheetName of sheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      continue
    }
    const addresses = new Set<string>()
    for (const { address, cell } of worksheetCellRecords(sheet)) {
      if (!hasImportableXlsxCellPayload(cell)) {
        continue
      }
      addresses.add(address)
      count += 1
      if (count > maxCandidateCount) {
        return { addressesBySheet: new Map(), count }
      }
    }
    if (addresses.size > 0) {
      addressesBySheet.set(sheetName, addresses)
    }
  }
  return { addressesBySheet, count }
}

function readImportedFillStyle(style: Record<string, unknown>): CellStyleRecord['fill'] | undefined {
  const fill = isRecord(style['fill']) ? style['fill'] : style
  if (fill['patternType'] !== 'solid') {
    return undefined
  }
  const backgroundColor = readRgbColor(fill['fgColor']) ?? readRgbColor(fill['bgColor'])
  return backgroundColor ? { backgroundColor } : undefined
}

function readImportedFontStyle(style: Record<string, unknown>): CellStyleFontSnapshot | undefined {
  const fontRecord = isRecord(style['font']) ? style['font'] : null
  if (!fontRecord) {
    return undefined
  }
  const font: CellStyleFontSnapshot = {}
  const family = typeof fontRecord['name'] === 'string' ? fontRecord['name'].trim() : ''
  if (family.length > 0) {
    font.family = family
  }
  const size = readFiniteNumber(fontRecord['sz']) ?? readFiniteNumber(fontRecord['size'])
  if (size !== null && size > 0) {
    font.size = size
  }
  if (fontRecord['bold'] === true) {
    font.bold = true
  }
  if (fontRecord['italic'] === true) {
    font.italic = true
  }
  if (fontRecord['underline'] === true || typeof fontRecord['underline'] === 'string') {
    font.underline = true
  }
  const color = readRgbColor(fontRecord['color'])
  if (color) {
    font.color = color
  }
  return Object.keys(font).length > 0 ? font : undefined
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
    case 'justify':
    case 'distributed':
      return value
    default:
      return undefined
  }
}

function readImportedAlignmentStyle(style: Record<string, unknown>): CellStyleAlignmentSnapshot | undefined {
  const alignmentRecord = isRecord(style['alignment']) ? style['alignment'] : null
  if (!alignmentRecord) {
    return undefined
  }
  const horizontal = readHorizontalAlignment(alignmentRecord['horizontal'])
  const vertical = readVerticalAlignment(alignmentRecord['vertical'])
  const indent = readImportedAlignmentNumber(alignmentRecord['indent'])
  const readingOrder = readImportedAlignmentNumber(alignmentRecord['readingOrder'])
  const textRotation = readImportedAlignmentNumber(alignmentRecord['textRotation'])
  const alignment: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['wrapText']) === true ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['shrinkToFit']) === true ? { shrinkToFit: true } : {}),
    ...(readingOrder !== null ? { readingOrder } : {}),
    ...(textRotation !== null ? { textRotation } : {}),
    ...(readImportedAlignmentBoolean(alignmentRecord['justifyLastLine']) === true ? { justifyLastLine: true } : {}),
  }
  return Object.keys(alignment).length > 0 ? alignment : undefined
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
    case 'mediumDashed':
    case 'dashDot':
    case 'dashDotDot':
    case 'slantDashDot':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return { style: 'dashed', weight: value === 'dashed' ? 'thin' : 'medium' }
    case 'dotted':
      return { style: 'dotted', weight: 'thin' }
    case 'double':
      return { style: 'double', weight: 'medium' }
    default:
      return null
  }
}

function readImportedBorderSide(value: unknown): CellBorderSideSnapshot | undefined {
  if (!isRecord(value)) {
    return undefined
  }
  const borderKind = readBorderKind(value['style'])
  if (!borderKind) {
    return undefined
  }
  return {
    ...borderKind,
    color: readRgbColor(value['color']) ?? '#000000',
  }
}

function readImportedBorderStyle(style: Record<string, unknown>): CellStyleBordersSnapshot | undefined {
  const borderRecord = isRecord(style['border']) ? style['border'] : null
  if (!borderRecord) {
    return undefined
  }
  const top = readImportedBorderSide(borderRecord['top'])
  const right = readImportedBorderSide(borderRecord['right'])
  const bottom = readImportedBorderSide(borderRecord['bottom'])
  const left = readImportedBorderSide(borderRecord['left'])
  const borders: CellStyleBordersSnapshot = {
    ...(top ? { top } : {}),
    ...(right ? { right } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
  }
  return Object.keys(borders).length > 0 ? borders : undefined
}

function readImportedProtectionStyle(style: Record<string, unknown>): CellStyleProtectionSnapshot | undefined {
  const protectionRecord = isRecord(style['protection']) ? style['protection'] : null
  if (!protectionRecord) {
    return undefined
  }
  return {
    ...(typeof protectionRecord['locked'] === 'boolean' ? { locked: protectionRecord['locked'] } : {}),
    ...(typeof protectionRecord['hidden'] === 'boolean' ? { hidden: protectionRecord['hidden'] } : {}),
  }
}

export function readImportedXlsxCellStyle(value: unknown): Omit<CellStyleRecord, 'id'> | null {
  if (!isRecord(value)) {
    return null
  }
  const fill = readImportedFillStyle(value)
  const font = readImportedFontStyle(value)
  const alignment = readImportedAlignmentStyle(value)
  const borders = readImportedBorderStyle(value)
  const protection = readImportedProtectionStyle(value)
  const style: Omit<CellStyleRecord, 'id'> = {
    ...(fill ? { fill } : {}),
    ...(font ? { font } : {}),
    ...(alignment ? { alignment } : {}),
    ...(borders ? { borders } : {}),
    ...(protection !== undefined ? { protection } : {}),
  }
  return Object.keys(style).length > 0 ? style : null
}

function internImportedStyle(style: Omit<CellStyleRecord, 'id'>, catalog: Map<string, CellStyleRecord>): string {
  const key = JSON.stringify(style)
  const existing = catalog.get(key)
  if (existing) {
    return existing.id
  }
  const record: CellStyleRecord = {
    id: `xlsx-style-${catalog.size + 1}`,
    ...style,
  }
  catalog.set(key, record)
  return record.id
}

function toUint8Array(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value)
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  return null
}

function readNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

function readImportedLiteralCellValue(cell: Record<string, unknown>): LiteralInput | undefined {
  if (cell['t'] === 'e') {
    if (Object.hasOwn(cell, 'w')) {
      const displayText = toLiteralInput(cell['w'])
      if (typeof displayText === 'string' && displayText.startsWith('#')) {
        return displayText
      }
    }
    if (!Object.hasOwn(cell, 'v')) {
      return undefined
    }
    const errorCode = cell['v']
    if (errorCode === undefined || errorCode === null) {
      return undefined
    }
    if (typeof errorCode === 'number') {
      return legacyExcelErrorTextByCode.get(errorCode) ?? '#ERROR!'
    }
    if (typeof errorCode === 'string' && errorCode.startsWith('#')) {
      return errorCode
    }
    return '#ERROR!'
  }
  return toLiteralInput(cell['v'])
}

function readImportedMacroCodeNames(workbook: XLSX.WorkBook): PreservedVbaProjectCodeNames {
  const workbookMetadata = isRecord(workbook.Workbook) ? workbook.Workbook : undefined
  const workbookProperties = isRecord(workbookMetadata?.['WBProps']) ? workbookMetadata['WBProps'] : undefined
  const workbookCodeName = readNonEmptyString(workbookProperties?.['CodeName'])
  const workbookSheets = workbookMetadata?.['Sheets']
  const sheetCodeNames = Array.isArray(workbookSheets)
    ? workbookSheets.flatMap((entry) => {
        if (!isRecord(entry)) {
          return []
        }
        const sheetName = readNonEmptyString(entry['name'])
        const codeName = readNonEmptyString(entry['CodeName'])
        return sheetName && codeName ? [{ sheetName, codeName }] : []
      })
    : []
  return {
    ...(workbookCodeName ? { workbookCodeName } : {}),
    ...(sheetCodeNames.length > 0 ? { sheetCodeNames } : {}),
  }
}

function addWorkbookWarnings(workbook: XLSX.WorkBook, warnings: string[], ignoredDefinedNameCount: number): void {
  if (workbook.vbaraw) {
    warnings.push('Macros were preserved but not executed during XLSX import.')
  }
  if (ignoredDefinedNameCount > 0) {
    warnings.push('Some defined names were ignored during XLSX import.')
  }
}

function readValidXlsxZipContainer(bytes: Uint8Array): Unzipped {
  try {
    return unzipSync(bytes)
  } catch {
    throw new InvalidXlsxZipContainerError()
  }
}

function importSheetJsWorkbook(
  data: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
): ImportedWorkbook {
  const parserData = workbookZip ? stripStyleOnlyBlankCellsForSheetJs(data, workbookZip) : data
  const workbook = XLSX.read(parserData, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellStyles: false,
    cellText: false,
    cellDates: false,
    bookFiles: true,
    bookVBA: true,
    dense: false,
  })
  const workbookName = normalizeWorkbookName(fileName)
  const warnings: string[] = []
  const importedDefinedNames = readImportedDefinedNames(workbook)
  addWorkbookWarnings(workbook, warnings, importedDefinedNames.ignoredCount)
  const styleCandidates = collectStyleCandidateAddresses(workbook, workbook.SheetNames, largeWorkbookStyleCandidateThreshold)
  const importedWorkbookStyles =
    styleCandidates.count === 0 || styleCandidates.count > largeWorkbookStyleCandidateThreshold
      ? new Map<string, Map<string, Omit<CellStyleRecord, 'id'>>>()
      : readImportedWorkbookFileStyles(workbook, workbook.SheetNames, {
          styleCandidateAddressesBySheet: styleCandidates.addressesBySheet,
        })
  const importedWorkbookSheetDimensions = readImportedWorkbookSheetDimensions(workbook, workbook.SheetNames)
  const importedWorkbookProperties = workbookZip ? readImportedWorkbookProperties(workbookZip) : undefined
  const importedCalculationSettings = workbookZip ? readImportedWorkbookCalculationSettings(workbookZip) : undefined
  const importedMacroPayload = toUint8Array(workbook.vbaraw)
  const importedMacroCodeNames = importedMacroPayload ? readImportedMacroCodeNames(workbook) : undefined
  const importedCharts = workbookZip ? readImportedWorkbookCharts(workbookZip, workbook.SheetNames) : undefined
  const importedPivots = workbookZip ? readImportedWorkbookPivots(workbookZip, workbook.SheetNames) : undefined
  const importedTables = workbookZip ? readImportedWorkbookTables(workbookZip, workbook.SheetNames) : undefined
  const importedFiltersBySheet = workbookZip ? readImportedWorkbookFilters(workbookZip, workbook.SheetNames) : new Map()
  const importedFreezePanesBySheet = workbookZip ? readImportedWorkbookFreezePanes(workbookZip, workbook.SheetNames) : new Map()
  const importedSheetTabColorsBySheet = workbookZip ? readImportedWorkbookSheetTabColors(workbookZip, workbook.SheetNames) : new Map()
  const importedSheetProtectionsBySheet = workbookZip ? readImportedWorkbookSheetProtections(workbookZip, workbook.SheetNames) : new Map()
  const importedProtectedRangesBySheet = workbookZip ? readImportedWorkbookProtectedRanges(workbookZip, workbook.SheetNames) : new Map()
  const importedSortsBySheet = workbookZip ? readImportedWorkbookSorts(workbookZip, workbook.SheetNames) : new Map()
  const importedValidationsBySheet = workbookZip ? readImportedWorkbookDataValidations(workbookZip, workbook.SheetNames) : new Map()
  const importedConditionalFormatsBySheet = workbookZip
    ? readImportedWorkbookConditionalFormats(workbookZip, workbook.SheetNames)
    : new Map()
  const importedExternalLinkCaches = workbookZip ? readImportedExternalLinkCaches(workbookZip) : new Map()

  let ignoredCommentsSeen = false
  const styleCatalog = new Map<string, CellStyleRecord>()
  const importedArrayFormulaSpills: NonNullable<WorkbookMetadataSnapshot['spills']> = []
  const previewSheets: ImportedWorkbookSheetPreview[] = []
  const sheets = workbook.SheetNames.map((sheetName, order) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet) {
      previewSheets.push(
        createSheetPreview({
          name: sheetName,
          rowCount: 0,
          columnCount: 0,
          nonEmptyCellCount: 0,
          readCellText: () => '',
        }),
      )
      return {
        id: order + 1,
        name: sheetName,
        order,
        cells: [],
      }
    }

    const importedComments = readImportedSheetComments(sheetName, sheet)
    if (importedComments.ignoredCount > 0 && !ignoredCommentsSeen) {
      ignoredCommentsSeen = true
      warnings.push('Some cell comments were ignored during XLSX import.')
    }
    const importedArrayFormulaSheetSpills = readImportedArrayFormulaSpills(sheetName, sheet)
    if (importedArrayFormulaSheetSpills) {
      importedArrayFormulaSpills.push(...importedArrayFormulaSheetSpills)
    }
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
    const styleRuns: RectangularStyleRun[] = []
    let openStyleRunsByKey = new Map<string, RectangularStyleRun>()
    let activeStyleRow: number | null = null
    let activeStyleRun: HorizontalStyleRun | null = null
    let activeStyleRowRuns: HorizontalStyleRun[] = []
    const importedStylesByAddress = importedWorkbookStyles.get(sheetName)
    const flushActiveStyleRun = () => {
      if (activeStyleRun) {
        activeStyleRowRuns.push(activeStyleRun)
        activeStyleRun = null
      }
    }
    const flushActiveStyleRow = () => {
      if (activeStyleRow === null) {
        return
      }
      flushActiveStyleRun()
      openStyleRunsByKey = mergeStyleRuns(activeStyleRow, activeStyleRowRuns, openStyleRunsByKey, styleRuns)
      activeStyleRowRuns = []
      activeStyleRow = null
    }
    const addStyleCell = (row: number, column: number, styleId: string) => {
      if (activeStyleRow === null) {
        activeStyleRow = row
      } else if (activeStyleRow !== row) {
        flushActiveStyleRow()
        activeStyleRow = row
      }
      if (activeStyleRun && activeStyleRun.styleId === styleId && activeStyleRun.endColumn + 1 === column) {
        activeStyleRun.endColumn = column
        return
      }
      flushActiveStyleRun()
      activeStyleRun = {
        styleId,
        startColumn: column,
        endColumn: column,
      }
    }
    const rowCount = range ? range.e.r + 1 : 0
    const columnCount = range ? range.e.c + 1 : 0
    const importableAddresses =
      styleCandidates.count <= largeWorkbookStyleCandidateThreshold ? styleCandidates.addressesBySheet.get(sheetName) : undefined
    const sheetCellEntries = range
      ? importableAddresses
        ? worksheetCellEntriesAtAddresses(sheet, importableAddresses)
        : worksheetCellEntries(sheet)
      : []
    for (const { address, cell, row, column } of sheetCellEntries) {
      const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { address }
      const formula = cell['f']
      if (typeof formula === 'string' && formula.trim().length > 0) {
        const externalReferenceFormula = translateImportedFormulaExternalReferences(formula, importedExternalLinkCaches).formula
        nextCell.formula = translateImportedFormulaStructuredReferences({
          formula: externalReferenceFormula,
          ownerSheetName: sheetName,
          ownerAddress: address,
          tables: importedTables,
        })
        const cachedLiteral = readImportedLiteralCellValue(cell)
        if (cachedLiteral !== undefined) {
          nextCell.value = cachedLiteral
        }
      } else {
        const literal = readImportedLiteralCellValue(cell)
        if (literal !== undefined) {
          nextCell.value = literal
        }
      }
      const importedFormat = readImportedNumberFormat(cell['z'])
      if (importedFormat !== undefined) {
        nextCell.format = importedFormat
      }
      const importedStyle = importedStylesByAddress?.get(address) ?? readImportedXlsxCellStyle(cell['s'])
      if (importedStyle) {
        addStyleCell(row, column, internImportedStyle(importedStyle, styleCatalog))
      } else if (activeStyleRow === row) {
        flushActiveStyleRun()
      }
      if (nextCell.value !== undefined || nextCell.formula !== undefined || nextCell.format !== undefined) {
        cells.push(nextCell)
      }
    }
    flushActiveStyleRow()
    const styleRanges = styleRunsToRanges(sheetName, styleRuns)

    previewSheets.push(
      createSheetPreview({
        name: sheetName,
        rowCount,
        columnCount,
        nonEmptyCellCount: cells.length,
        readCellText: (row, col) => {
          const cell = worksheetCellAt(sheet, row, col)
          if (!cell) {
            return ''
          }
          const formula = cell['f']
          if (typeof formula === 'string' && formula.trim().length > 0) {
            return `=${formula}`
          }
          return toDisplayText(readImportedLiteralCellValue(cell))
        },
      }),
    )

    const importedSheetDimensions = importedWorkbookSheetDimensions.get(sheetName)
    const rows = importedSheetDimensions?.rows ?? buildRowEntries(sheet['!rows'])
    const columns = importedSheetDimensions ? importedSheetDimensions.columns : buildColumnEntries(sheet['!cols'])
    const importedFreezePane = importedFreezePanesBySheet.get(sheetName)
    const importedSheetTabColor = importedSheetTabColorsBySheet.get(sheetName)
    const merges = buildMergeEntries(sheetName, sheet['!merges'])
    const importedSheetProtection = importedSheetProtectionsBySheet.get(sheetName)
    const importedProtectedRanges = importedProtectedRangesBySheet.get(sheetName)
    const importedSorts = importedSortsBySheet.get(sheetName)
    const importedFilters = importedFiltersBySheet.get(sheetName)
    const importedValidations = importedValidationsBySheet.get(sheetName)
    const importedConditionalFormats = importedConditionalFormatsBySheet.get(sheetName)
    const metadata =
      rows ||
      columns ||
      styleRanges.length > 0 ||
      importedFreezePane ||
      importedSheetTabColor ||
      merges ||
      importedSheetProtection ||
      importedProtectedRanges ||
      importedSorts ||
      importedFilters ||
      importedValidations ||
      importedConditionalFormats ||
      importedComments.commentThreads
        ? {
            ...(rows ? { rows } : {}),
            ...(columns ? { columns } : {}),
            ...(styleRanges.length > 0 ? { styleRanges } : {}),
            ...(importedFreezePane ? { freezePane: importedFreezePane } : {}),
            ...(importedSheetTabColor ? { tabColor: importedSheetTabColor } : {}),
            ...(merges ? { merges } : {}),
            ...(importedSheetProtection ? { sheetProtection: importedSheetProtection } : {}),
            ...(importedProtectedRanges ? { protectedRanges: importedProtectedRanges } : {}),
            ...(importedSorts ? { sorts: importedSorts } : {}),
            ...(importedFilters ? { filters: importedFilters } : {}),
            ...(importedValidations ? { validations: importedValidations } : {}),
            ...(importedConditionalFormats ? { conditionalFormats: importedConditionalFormats } : {}),
            ...(importedComments.commentThreads ? { commentThreads: importedComments.commentThreads } : {}),
          }
        : undefined

    return {
      id: order + 1,
      name: sheetName,
      order,
      ...(metadata ? { metadata } : {}),
      cells,
    }
  })

  const workbookMetadata: WorkbookMetadataSnapshot = {
    ...(importedWorkbookProperties ? { properties: importedWorkbookProperties } : {}),
    ...(importedCalculationSettings ? { calculationSettings: importedCalculationSettings } : {}),
    ...(importedMacroPayload ? { macroPayloads: [createPreservedVbaProjectPayload(importedMacroPayload, importedMacroCodeNames)] } : {}),
    ...(styleCatalog.size > 0 ? { styles: [...styleCatalog.values()] } : {}),
    ...(importedDefinedNames.definedNames ? { definedNames: importedDefinedNames.definedNames } : {}),
    ...(importedTables ? { tables: importedTables } : {}),
    ...(importedArrayFormulaSpills.length > 0 ? { spills: importedArrayFormulaSpills } : {}),
    ...(importedPivots ? { pivots: importedPivots } : {}),
    ...(importedCharts ? { charts: importedCharts } : {}),
  }

  return {
    snapshot: {
      version: 1,
      workbook: {
        name: workbookName,
        ...(Object.keys(workbookMetadata).length > 0 ? { metadata: workbookMetadata } : {}),
      },
      sheets,
    },
    workbookName,
    sheetNames: workbook.SheetNames,
    warnings,
    preview: createWorkbookPreview({
      contentType,
      fileName,
      fileSizeBytes: data.byteLength,
      workbookName,
      sheets: previewSheets,
      warnings,
    }),
  }
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const workbookZip = readValidXlsxZipContainer(data)
  return importSheetJsWorkbook(data, fileName, XLSX_CONTENT_TYPE, workbookZip)
}

export function importXlsb(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return importSheetJsWorkbook(data, fileName, XLSB_CONTENT_TYPE, null)
}

export function importCsv(text: string, fileName: string, options: CsvImportOptions = {}): ImportedWorkbook {
  const workbookName = normalizeWorkbookName(fileName)
  const sheetName = normalizeCsvSheetName(workbookName)
  const csvOptions = resolveCsvParseOptions(text, options)
  const rows = parseCsv(text, csvOptions)
  const textColumnIndexes = inferCsvTextColumnIndexes(rows)
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  let nonEmptyCellCount = 0
  let hasRaggedRows = false
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0)

  rows.forEach((row, rowIndex) => {
    if (row.length !== columnCount) {
      hasRaggedRows = true
    }
    row.forEach((raw, colIndex) => {
      const parsed =
        textColumnIndexes.has(colIndex) && rowIndex > 0 && raw.trim() !== '' ? { value: raw } : parseCsvCellInput(raw, csvOptions)
      if (!parsed) {
        return
      }
      nonEmptyCellCount += 1
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
      if (parsed.formula !== undefined) {
        cells.push({ address, formula: parsed.formula })
        return
      }
      if (parsed.value !== undefined) {
        cells.push({ address, value: parsed.value })
      }
    })
  })

  const warnings = hasRaggedRows ? ['CSV rows had inconsistent column counts. Missing cells were treated as blanks.'] : []
  const previewSheet = createSheetPreview({
    name: sheetName,
    rowCount: rows.length,
    columnCount,
    nonEmptyCellCount,
    readCellText: (row, col) => rows[row]?.[col] ?? '',
  })

  return {
    snapshot: {
      version: 1,
      workbook: {
        name: workbookName,
      },
      sheets: [
        {
          id: 1,
          name: sheetName,
          order: 0,
          cells,
        },
      ],
    },
    workbookName,
    sheetNames: [sheetName],
    warnings,
    preview: createWorkbookPreview({
      contentType: CSV_CONTENT_TYPE,
      fileName,
      fileSizeBytes: new TextEncoder().encode(text).byteLength,
      workbookName,
      sheets: [previewSheet],
      warnings,
    }),
  }
}

export function importWorkbookFile(
  bytes: Uint8Array | ArrayBuffer,
  fileName: string,
  contentType: string,
  options: WorkbookImportFileOptions = {},
): ImportedWorkbook {
  const normalizedContentType = normalizeWorkbookImportContentType(contentType)
  if (normalizedContentType === XLSX_CONTENT_TYPE) {
    return importXlsx(bytes, fileName)
  }
  if (normalizedContentType === XLSB_CONTENT_TYPE) {
    return importXlsb(bytes, fileName)
  }
  if (normalizedContentType === CSV_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return importCsv(new TextDecoder().decode(data), fileName, options.csv)
  }
  throw new Error('Unsupported workbook import content type')
}

function inferCsvTextColumnIndexes(rows: readonly (readonly string[])[]): Set<number> {
  const header = rows[0]
  const textColumnIndexes = new Set<number>()
  if (!header) {
    return textColumnIndexes
  }

  header.forEach((rawHeader, colIndex) => {
    const headerText = rawHeader.trim().toLowerCase().replaceAll('_', ' ').replaceAll('-', ' ')
    if (isIdentifierLikeCsvHeader(headerText)) {
      textColumnIndexes.add(colIndex)
    }
  })
  return textColumnIndexes
}

function isIdentifierLikeCsvHeader(headerText: string): boolean {
  return /^(?:account|acct|id|code|sku)(?: (?:id|number|no|code))?$/u.test(headerText)
}
