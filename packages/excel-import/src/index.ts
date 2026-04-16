import * as XLSX from 'xlsx'

import { CSV_CONTENT_TYPE, XLSX_CONTENT_TYPE, type WorkbookImportContentType } from '@bilig/agent-api'
import { parseCsv, parseCsvCellInput } from '@bilig/core'
import type { WorkbookAxisEntrySnapshot, WorkbookSnapshot } from '@bilig/protocol'

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
  return trimmed.replace(/\.(xlsx|csv)$/i, '') || 'Imported workbook'
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

function addWorkbookWarnings(workbook: XLSX.WorkBook, warnings: string[]): void {
  if (workbook.vbaraw) {
    warnings.push('Macros were ignored during XLSX import.')
  }
  const definedNames = workbook.Workbook?.Names
  if (Array.isArray(definedNames) && definedNames.length > 0) {
    warnings.push('Defined names were ignored during XLSX import.')
  }
}

function addSheetWarnings(sheetName: string, sheet: XLSX.WorkSheet, warnings: string[], ignoredComments: { seen: boolean }): void {
  const merges = sheet['!merges']
  if (Array.isArray(merges) && merges.length > 0) {
    warnings.push(`Merged cells on ${sheetName} were ignored during XLSX import.`)
  }
  Object.values(sheet).forEach((value) => {
    if (!isRecord(value)) {
      return
    }
    if (!ignoredComments.seen && Array.isArray(value['c']) && value['c'].length > 0) {
      ignoredComments.seen = true
      warnings.push('Cell comments were ignored during XLSX import.')
    }
  })
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
  })
  const workbookName = normalizeWorkbookName(fileName)
  const warnings: string[] = []
  addWorkbookWarnings(workbook, warnings)

  const ignoredComments = { seen: false }
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

    addSheetWarnings(sheetName, sheet, warnings, ignoredComments)
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
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
          if (typeof cell.z === 'string' && cell.z.trim().length > 0) {
            nextCell.format = cell.z
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
    const metadata =
      rows || columns
        ? {
            ...(rows ? { rows } : {}),
            ...(columns ? { columns } : {}),
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

  return {
    snapshot: {
      version: 1,
      workbook: {
        name: workbookName,
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
