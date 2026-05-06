import * as XLSX from 'xlsx'
import * as XLSXStyle from 'xlsx-js-style'

import type {
  LiteralInput,
  WorkbookAxisEntrySnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { addExportCalculationSettingsToXlsxBytes } from './xlsx-calculation-settings.js'
import { addExportChartsToXlsxBytes } from './xlsx-charts.js'
import { addExportCommentsToWorksheet } from './xlsx-comments.js'
import { addExportConditionalFormatsToXlsxBytes } from './xlsx-conditional-formats.js'
import { buildExportDefinedNames } from './xlsx-defined-names.js'
import { addExportFiltersToXlsxBytes } from './xlsx-filters.js'
import { addExportFreezePanesToXlsxBytes } from './xlsx-freeze-panes.js'
import { addExportPivotsToXlsxBytes } from './xlsx-pivots.js'
import { addExportProtectedRangesToXlsxBytes } from './xlsx-protected-ranges.js'
import { addExportSheetProtectionsToXlsxBytes } from './xlsx-sheet-protection.js'
import { addExportSortsToXlsxBytes } from './xlsx-sorts.js'
import { addExportStylesToWorksheet } from './xlsx-styles.js'
import { addExportTablesToXlsxBytes } from './xlsx-tables.js'
import { addExportDataValidationsToXlsxBytes } from './xlsx-validations.js'
import { addExportWorkbookPropertiesToXlsxBytes } from './xlsx-workbook-properties.js'
import { decodePreservedVbaProjectPayload } from './xlsx-macros.js'

function buildExportColumns(columns: readonly WorkbookAxisEntrySnapshot[] | undefined): XLSX.ColInfo[] | undefined {
  if (!columns || columns.length === 0) {
    return undefined
  }
  const maxIndex = columns.reduce((max, column) => Math.max(max, column.index), -1)
  if (maxIndex < 0) {
    return undefined
  }
  const output = Array.from({ length: maxIndex + 1 }, (): XLSX.ColInfo => ({}))
  for (const column of columns) {
    const target = output[column.index]
    if (!target) {
      continue
    }
    if (typeof column.size === 'number' && Number.isFinite(column.size) && column.size > 0) {
      target.wpx = column.size
    }
    if (column.hidden === true) {
      target.hidden = true
    }
  }
  return output.some((column) => Object.keys(column).length > 0) ? output : undefined
}

function buildExportRows(rows: readonly WorkbookAxisEntrySnapshot[] | undefined): XLSX.RowInfo[] | undefined {
  if (!rows || rows.length === 0) {
    return undefined
  }
  const maxIndex = rows.reduce((max, row) => Math.max(max, row.index), -1)
  if (maxIndex < 0) {
    return undefined
  }
  const output = Array.from({ length: maxIndex + 1 }, (): XLSX.RowInfo => ({}))
  for (const row of rows) {
    const target = output[row.index]
    if (!target) {
      continue
    }
    if (typeof row.size === 'number' && Number.isFinite(row.size) && row.size > 0) {
      target.hpx = row.size
    }
    if (row.hidden === true) {
      target.hidden = true
    }
  }
  return output.some((row) => Object.keys(row).length > 0) ? output : undefined
}

function buildExportMerges(merges: readonly WorkbookMergeRangeSnapshot[] | undefined): XLSX.Range[] | undefined {
  if (!merges || merges.length === 0) {
    return undefined
  }
  return merges.map((merge) => XLSX.utils.decode_range(`${merge.startAddress}:${merge.endAddress}`))
}

function updateWorksheetBounds(bounds: XLSX.Range | null, address: string): XLSX.Range {
  const decoded = XLSX.utils.decode_cell(address)
  if (!bounds) {
    return {
      s: { r: decoded.r, c: decoded.c },
      e: { r: decoded.r, c: decoded.c },
    }
  }
  return {
    s: {
      r: Math.min(bounds.s.r, decoded.r),
      c: Math.min(bounds.s.c, decoded.c),
    },
    e: {
      r: Math.max(bounds.e.r, decoded.r),
      c: Math.max(bounds.e.c, decoded.c),
    },
  }
}

function updateWorksheetBoundsForAxis(bounds: XLSX.Range | null, axis: 'row' | 'column', index: number): XLSX.Range {
  const address = axis === 'row' ? XLSX.utils.encode_cell({ r: index, c: 0 }) : XLSX.utils.encode_cell({ r: 0, c: index })
  return updateWorksheetBounds(bounds, address)
}

function inferExportWorksheetRange(sheet: WorkbookSnapshot['sheets'][number]): string | undefined {
  let bounds: XLSX.Range | null = null
  for (const cell of sheet.cells) {
    bounds = updateWorksheetBounds(bounds, cell.address)
  }
  for (const merge of sheet.metadata?.merges ?? []) {
    bounds = updateWorksheetBounds(bounds, merge.startAddress)
    bounds = updateWorksheetBounds(bounds, merge.endAddress)
  }
  for (const thread of sheet.metadata?.commentThreads ?? []) {
    bounds = updateWorksheetBounds(bounds, thread.address)
  }
  for (const styleRange of sheet.metadata?.styleRanges ?? []) {
    bounds = updateWorksheetBounds(bounds, styleRange.range.startAddress)
    bounds = updateWorksheetBounds(bounds, styleRange.range.endAddress)
  }
  for (const row of sheet.metadata?.rows ?? []) {
    bounds = updateWorksheetBoundsForAxis(bounds, 'row', row.index)
  }
  for (const column of sheet.metadata?.columns ?? []) {
    bounds = updateWorksheetBoundsForAxis(bounds, 'column', column.index)
  }
  return bounds ? XLSX.utils.encode_range(bounds) : undefined
}

function cellTypeForLiteral(value: LiteralInput): XLSX.ExcelDataType | undefined {
  if (typeof value === 'number') {
    return 'n'
  }
  if (typeof value === 'boolean') {
    return 'b'
  }
  if (typeof value === 'string') {
    return 's'
  }
  return undefined
}

function buildExportCell(cell: WorkbookSnapshot['sheets'][number]['cells'][number]): XLSXStyle.CellObject | null {
  const output: XLSXStyle.CellObject = { t: 'z' }
  if (cell.value !== undefined && cell.value !== null) {
    const type = cellTypeForLiteral(cell.value)
    if (type) {
      output.t = type
      output.v = cell.value
    }
  }
  if (typeof cell.formula === 'string' && cell.formula.trim().length > 0) {
    output.f = cell.formula.trim().replace(/^=/, '')
    output.t = output.t ?? 'n'
  }
  if (typeof cell.format === 'string' && cell.format.trim().length > 0) {
    output.z = cell.format
  }
  return output.v !== undefined || output.f !== undefined || output.z !== undefined ? output : null
}

const invalidExportSheetNameCharacters = ['[', ']', ':', '*', '?', '/', '\\'] as const

function normalizeExportSheetName(name: string, order: number, usedNames: Set<string>): string {
  let sanitized = name
  for (const character of invalidExportSheetNameCharacters) {
    sanitized = sanitized.split(character).join(' ')
  }
  const baseName = sanitized.trim() || `Sheet${order + 1}`
  let candidate = baseName.slice(0, 31)
  let suffix = 1
  while (usedNames.has(candidate)) {
    const suffixText = ` ${String(suffix)}`
    candidate = `${baseName.slice(0, 31 - suffixText.length)}${suffixText}`
    suffix += 1
  }
  usedNames.add(candidate)
  return candidate
}

function toUint8Array(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value)
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  throw new Error('XLSX writer returned unsupported output bytes')
}

function applyMacroCodeNamesToWorkbook(
  workbook: XLSXStyle.WorkBook,
  macroPayload: WorkbookMacroPayloadSnapshot | undefined,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): void {
  if (!macroPayload?.workbookCodeName && (!macroPayload?.sheetCodeNames || macroPayload.sheetCodeNames.length === 0)) {
    return
  }

  const workbookMetadata = {
    ...workbook.Workbook,
  }
  if (macroPayload.workbookCodeName) {
    workbookMetadata.WBProps = {
      ...workbookMetadata.WBProps,
      CodeName: macroPayload.workbookCodeName,
    }
  }

  if (macroPayload.sheetCodeNames && macroPayload.sheetCodeNames.length > 0) {
    const codeNamesByExportSheetName = new Map<string, string>()
    for (const entry of macroPayload.sheetCodeNames) {
      const exportSheetName = exportSheetNamesByOriginalName.get(entry.sheetName) ?? entry.sheetName
      codeNamesByExportSheetName.set(exportSheetName, entry.codeName)
    }
    const existingSheets = workbookMetadata.Sheets ?? []
    workbookMetadata.Sheets = workbook.SheetNames.map((sheetName, index) => {
      const codeName = codeNamesByExportSheetName.get(sheetName) ?? existingSheets[index]?.CodeName
      return {
        ...existingSheets[index],
        name: sheetName,
        ...(codeName ? { CodeName: codeName } : {}),
      }
    })
  }

  workbook.Workbook = workbookMetadata
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  const workbook = XLSXStyle.utils.book_new()
  const usedNames = new Set<string>()
  const exportSheetNamesByOriginalName = new Map<string, string>()

  for (const sheet of snapshot.sheets.toSorted((left, right) => left.order - right.order)) {
    const worksheet: XLSXStyle.WorkSheet = {}
    for (const cell of sheet.cells) {
      const exportCell = buildExportCell(cell)
      if (exportCell) {
        worksheet[cell.address] = exportCell
      }
    }

    const ref = inferExportWorksheetRange(sheet)
    if (ref) {
      worksheet['!ref'] = ref
    }
    const columns = buildExportColumns(sheet.metadata?.columns)
    if (columns) {
      worksheet['!cols'] = columns
    }
    const rows = buildExportRows(sheet.metadata?.rows)
    if (rows) {
      worksheet['!rows'] = rows
    }
    const merges = buildExportMerges(sheet.metadata?.merges)
    if (merges) {
      worksheet['!merges'] = merges
    }
    addExportStylesToWorksheet(worksheet, sheet.metadata?.styleRanges, snapshot.workbook.metadata?.styles)
    addExportCommentsToWorksheet(worksheet, sheet.metadata?.commentThreads)

    const exportSheetName = normalizeExportSheetName(sheet.name, sheet.order, usedNames)
    exportSheetNamesByOriginalName.set(sheet.name, exportSheetName)
    XLSXStyle.utils.book_append_sheet(workbook, worksheet, exportSheetName)
  }

  const definedNames = buildExportDefinedNames(snapshot.workbook.metadata?.definedNames, exportSheetNamesByOriginalName)
  if (definedNames) {
    workbook.Workbook = {
      ...workbook.Workbook,
      Names: definedNames,
    }
  }

  const macroPayload = snapshot.workbook.metadata?.macroPayloads?.[0]
  const preservedVbaProject = decodePreservedVbaProjectPayload(macroPayload)
  if (preservedVbaProject) {
    const macroWorkbook = workbook as { vbaraw?: Uint8Array }
    macroWorkbook.vbaraw = preservedVbaProject
    applyMacroCodeNamesToWorkbook(workbook, macroPayload, exportSheetNamesByOriginalName)
  }

  const bytes = toUint8Array(
    XLSXStyle.write(workbook, {
      bookType: preservedVbaProject ? 'xlsm' : 'xlsx',
      type: 'buffer',
      cellStyles: true,
      bookVBA: Boolean(preservedVbaProject),
    }) as unknown,
  )
  return addExportChartsToXlsxBytes(
    addExportPivotsToXlsxBytes(
      addExportTablesToXlsxBytes(
        addExportDataValidationsToXlsxBytes(
          addExportConditionalFormatsToXlsxBytes(
            addExportSortsToXlsxBytes(
              addExportFiltersToXlsxBytes(
                addExportProtectedRangesToXlsxBytes(
                  addExportSheetProtectionsToXlsxBytes(
                    addExportFreezePanesToXlsxBytes(
                      addExportCalculationSettingsToXlsxBytes(addExportWorkbookPropertiesToXlsxBytes(bytes, snapshot), snapshot),
                      snapshot,
                    ),
                    snapshot,
                  ),
                  snapshot,
                ),
                snapshot,
              ),
              snapshot,
            ),
            snapshot,
          ),
          snapshot,
          exportSheetNamesByOriginalName,
        ),
        snapshot,
        exportSheetNamesByOriginalName,
      ),
      snapshot,
      exportSheetNamesByOriginalName,
    ),
    snapshot,
    exportSheetNamesByOriginalName,
  )
}
