import * as XLSX from 'xlsx'

import { CSV_CONTENT_TYPE, XLSX_CONTENT_TYPE, type WorkbookImportContentType } from '@bilig/agent-api'
import { parseCsv, parseCsvCellInput } from '@bilig/core'
import type {
  CellBorderSideSnapshot,
  CellBorderStyle,
  CellBorderWeight,
  CellHorizontalAlignment,
  CellStyleAlignmentSnapshot,
  CellStyleBordersSnapshot,
  CellStyleFontSnapshot,
  CellStyleRecord,
  CellVerticalAlignment,
  SheetStyleRangeSnapshot,
  WorkbookAxisEntrySnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookMetadataSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readImportedWorkbookCalculationSettings } from './xlsx-calculation-settings.js'
import { readImportedWorkbookCharts } from './xlsx-charts.js'
import { readImportedSheetComments } from './xlsx-comments.js'
import { readImportedWorkbookConditionalFormats } from './xlsx-conditional-formats.js'
import { readImportedDefinedNames } from './xlsx-defined-names.js'
import { readImportedWorkbookFilters } from './xlsx-filters.js'
import { readImportedWorkbookFreezePanes } from './xlsx-freeze-panes.js'
import { readImportedWorkbookPivots } from './xlsx-pivots.js'
import { readImportedWorkbookProtectedRanges } from './xlsx-protected-ranges.js'
import { readImportedWorkbookSheetProtections } from './xlsx-sheet-protection.js'
import { readImportedWorkbookSorts } from './xlsx-sorts.js'
import { readImportedWorkbookFileStyles } from './xlsx-styles.js'
import { readImportedWorkbookTables } from './xlsx-tables.js'
import { readImportedWorkbookDataValidations } from './xlsx-validations.js'
import { readImportedWorkbookProperties } from './xlsx-workbook-properties.js'
import { createPreservedVbaProjectPayload, type PreservedVbaProjectCodeNames } from './xlsx-macros.js'

export { exportXlsx } from './xlsx-export.js'

const PREVIEW_ROW_LIMIT = 8
const PREVIEW_COLUMN_LIMIT = 6

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

export interface ImportedWorkbookSheetPreview {
  name: string
  rowCount: number
  columnCount: number
  nonEmptyCellCount: number
  previewRows: readonly (readonly string[])[]
}

interface SheetColumnInfo {
  index: number
  size: number
}

interface SheetRowInfo {
  index: number
  size: number
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function normalizeWorkbookName(fileName: string): string {
  const trimmed = fileName.trim()
  if (trimmed.length === 0) {
    return 'Imported workbook'
  }
  return trimmed.replace(/\.(xlsx|xlsm|csv)$/i, '') || 'Imported workbook'
}

function normalizeCsvSheetName(workbookName: string): string {
  const trimmed = workbookName.trim()
  return trimmed.length > 0 ? trimmed : 'Sheet1'
}

function toLiteralInput(value: unknown) {
  if (value === null || value === undefined) {
    return undefined
  }
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) {
    return value.getTime()
  }
  return undefined
}

function toDisplayText(value: string | number | boolean | undefined): string {
  if (value === undefined) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE'
  }
  return String(value)
}

function createSheetPreview(input: {
  name: string
  rowCount: number
  columnCount: number
  nonEmptyCellCount: number
  readCellText: (row: number, col: number) => string
}): ImportedWorkbookSheetPreview {
  const previewRows: string[][] = []
  const previewRowCount = Math.min(input.rowCount, PREVIEW_ROW_LIMIT)
  const previewColumnCount = Math.min(input.columnCount, PREVIEW_COLUMN_LIMIT)
  for (let row = 0; row < previewRowCount; row += 1) {
    const values: string[] = []
    for (let col = 0; col < previewColumnCount; col += 1) {
      values.push(input.readCellText(row, col))
    }
    previewRows.push(values)
  }
  return {
    name: input.name,
    rowCount: input.rowCount,
    columnCount: input.columnCount,
    nonEmptyCellCount: input.nonEmptyCellCount,
    previewRows,
  }
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
        ? Math.round(entry['wpx'])
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
        ? Math.round(entry['hpx'])
        : typeof entry['hpt'] === 'number'
          ? toPixelSize(entry['hpt'], 'pt')
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
    id: `row:${index}`,
    index,
    size,
  }))
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
  const indent = readFiniteNumber(alignmentRecord['indent'])
  const alignment: CellStyleAlignmentSnapshot = {
    ...(horizontal ? { horizontal } : {}),
    ...(vertical ? { vertical } : {}),
    ...(alignmentRecord['wrapText'] === true ? { wrap: true } : {}),
    ...(indent !== null && indent >= 0 ? { indent } : {}),
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

export function readImportedXlsxCellStyle(value: unknown): Omit<CellStyleRecord, 'id'> | null {
  if (!isRecord(value)) {
    return null
  }
  const fill = readImportedFillStyle(value)
  const font = readImportedFontStyle(value)
  const alignment = readImportedAlignmentStyle(value)
  const borders = readImportedBorderStyle(value)
  const style: Omit<CellStyleRecord, 'id'> = {
    ...(fill ? { fill } : {}),
    ...(font ? { font } : {}),
    ...(alignment ? { alignment } : {}),
    ...(borders ? { borders } : {}),
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

function buildMergeEntries(sheetName: string, merges: readonly XLSX.Range[] | undefined): WorkbookMergeRangeSnapshot[] | undefined {
  if (!Array.isArray(merges) || merges.length === 0) {
    return undefined
  }
  return merges.map((range) => ({
    sheetName,
    startAddress: XLSX.utils.encode_cell(range.s),
    endAddress: XLSX.utils.encode_cell(range.e),
  }))
}

function createCellRange(sheetName: string, address: string) {
  return {
    sheetName,
    startAddress: address,
    endAddress: address,
  }
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

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const workbook = XLSX.read(data, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellText: false,
    cellDates: false,
    bookFiles: true,
    bookVBA: true,
  })
  const workbookName = normalizeWorkbookName(fileName)
  const warnings: string[] = []
  const importedDefinedNames = readImportedDefinedNames(workbook)
  addWorkbookWarnings(workbook, warnings, importedDefinedNames.ignoredCount)
  const importedWorkbookStyles = readImportedWorkbookFileStyles(workbook, workbook.SheetNames)
  const importedWorkbookProperties = readImportedWorkbookProperties(data)
  const importedCalculationSettings = readImportedWorkbookCalculationSettings(data)
  const importedMacroPayload = toUint8Array(workbook.vbaraw)
  const importedMacroCodeNames = importedMacroPayload ? readImportedMacroCodeNames(workbook) : undefined
  const importedCharts = readImportedWorkbookCharts(data, workbook.SheetNames)
  const importedPivots = readImportedWorkbookPivots(data, workbook.SheetNames)
  const importedTables = readImportedWorkbookTables(data, workbook.SheetNames)
  const importedFiltersBySheet = readImportedWorkbookFilters(data, workbook.SheetNames)
  const importedFreezePanesBySheet = readImportedWorkbookFreezePanes(data, workbook.SheetNames)
  const importedSheetProtectionsBySheet = readImportedWorkbookSheetProtections(data, workbook.SheetNames)
  const importedProtectedRangesBySheet = readImportedWorkbookProtectedRanges(data, workbook.SheetNames)
  const importedSortsBySheet = readImportedWorkbookSorts(data, workbook.SheetNames)
  const importedValidationsBySheet = readImportedWorkbookDataValidations(data, workbook.SheetNames)
  const importedConditionalFormatsBySheet = readImportedWorkbookConditionalFormats(data, workbook.SheetNames)

  let ignoredCommentsSeen = false
  const styleCatalog = new Map<string, CellStyleRecord>()
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
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
    const styleRanges: SheetStyleRangeSnapshot[] = []
    const rowCount = range ? range.e.r + 1 : 0
    const columnCount = range ? range.e.c + 1 : 0
    if (range) {
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let col = range.s.c; col <= range.e.c; col += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: col })
          const cell = sheet[address]
          if (!cell) {
            continue
          }
          const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { address }
          if (typeof cell.f === 'string' && cell.f.trim().length > 0) {
            nextCell.formula = cell.f
          } else {
            const literal = toLiteralInput(cell.v)
            if (literal !== undefined) {
              nextCell.value = literal
            }
          }
          const importedFormat = readImportedNumberFormat(cell.z)
          if (importedFormat !== undefined) {
            nextCell.format = importedFormat
          }
          const importedStyle = importedWorkbookStyles.get(sheetName)?.get(address) ?? readImportedXlsxCellStyle(cell.s)
          if (importedStyle) {
            styleRanges.push({
              range: createCellRange(sheetName, address),
              styleId: internImportedStyle(importedStyle, styleCatalog),
            })
          }
          if (nextCell.value !== undefined || nextCell.formula !== undefined || nextCell.format !== undefined) {
            cells.push(nextCell)
          }
        }
      }
    }

    previewSheets.push(
      createSheetPreview({
        name: sheetName,
        rowCount,
        columnCount,
        nonEmptyCellCount: cells.length,
        readCellText: (row, col) => {
          const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })]
          if (!cell) {
            return ''
          }
          if (typeof cell.f === 'string' && cell.f.trim().length > 0) {
            return `=${cell.f}`
          }
          return toDisplayText(toLiteralInput(cell.v))
        },
      }),
    )

    const rows = buildRowEntries(sheet['!rows'])
    const columns = buildColumnEntries(sheet['!cols'])
    const importedFreezePane = importedFreezePanesBySheet.get(sheetName)
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
      contentType: XLSX_CONTENT_TYPE,
      fileName,
      fileSizeBytes: data.byteLength,
      workbookName,
      sheets: previewSheets,
      warnings,
    }),
  }
}

export function importCsv(text: string, fileName: string): ImportedWorkbook {
  const workbookName = normalizeWorkbookName(fileName)
  const sheetName = normalizeCsvSheetName(workbookName)
  const rows = parseCsv(text)
  const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
  let nonEmptyCellCount = 0
  let hasRaggedRows = false
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0)

  rows.forEach((row, rowIndex) => {
    if (row.length !== columnCount) {
      hasRaggedRows = true
    }
    row.forEach((raw, colIndex) => {
      const parsed = parseCsvCellInput(raw)
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
  contentType: WorkbookImportContentType,
): ImportedWorkbook {
  if (contentType === XLSX_CONTENT_TYPE) {
    return importXlsx(bytes, fileName)
  }
  if (contentType === CSV_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return importCsv(new TextDecoder().decode(data), fileName)
  }
  throw new Error('Unsupported workbook import content type')
}
