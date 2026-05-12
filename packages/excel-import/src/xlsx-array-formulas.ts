import { unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type {
  WorkbookSheetArrayFormulaSnapshot,
  WorkbookSheetArrayFormulasSnapshot,
  WorkbookSnapshot,
  WorkbookSpillSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import { escapeXml, setZipText } from './xlsx-pivot-artifacts.js'

const cellElementPattern = /<c\b(?<attributes>[^>]*)>(?<body>[\s\S]*?)<\/c>/gu
const formulaElementPattern = /<f\b[^>]*(?:\/>|>[\s\S]*?<\/f>)/u
const formulaElementGlobalPattern = /<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/gu

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function isWorksheetCellAddress(value: string): boolean {
  return /^[A-Z]{1,3}[1-9][0-9]*$/u.test(value)
}

function cellColumnIndex(address: string): number {
  try {
    return XLSX.utils.decode_cell(address).c
  } catch {
    return Number.MAX_SAFE_INTEGER
  }
}

function cellRowNumber(address: string): number | null {
  const match = /^[A-Z]+([1-9][0-9]*)$/u.exec(address)
  return match ? Number(match[1]) : null
}

function decodeArrayFormulaRange(value: string): XLSX.Range | undefined {
  try {
    return XLSX.utils.decode_range(value)
  } catch {
    return undefined
  }
}

export function readImportedArrayFormulaSpills(sheetName: string, sheet: XLSX.WorkSheet): WorkbookSpillSnapshot[] | undefined {
  const spills: WorkbookSpillSnapshot[] = []
  for (const address in sheet) {
    const cell: unknown = sheet[address]
    if (!isWorksheetCellAddress(address) || !isRecord(cell)) {
      continue
    }
    const formula = cell['f']
    const arrayRangeText = cell['F']
    if (typeof formula !== 'string' || formula.trim().length === 0 || typeof arrayRangeText !== 'string') {
      continue
    }
    const range = decodeArrayFormulaRange(arrayRangeText.trim())
    if (!range) {
      continue
    }
    const owner = XLSX.utils.decode_cell(address)
    if (range.s.r !== owner.r || range.s.c !== owner.c) {
      continue
    }
    const rows = range.e.r - range.s.r + 1
    const cols = range.e.c - range.s.c + 1
    if (rows <= 1 && cols <= 1) {
      continue
    }
    spills.push({
      sheetName,
      address: XLSX.utils.encode_cell(range.s),
      rows,
      cols,
    })
  }
  return spills.length > 0 ? spills : undefined
}

function readArrayFormulaSnapshots(sheetXml: string | null): WorkbookSheetArrayFormulaSnapshot[] {
  if (!sheetXml) {
    return []
  }
  const formulas: WorkbookSheetArrayFormulaSnapshot[] = []
  cellElementPattern.lastIndex = 0
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const attributes = match.groups?.['attributes'] ?? ''
    const body = match.groups?.['body'] ?? ''
    const address = readAttribute(attributes, 'r')
    formulaElementGlobalPattern.lastIndex = 0
    const formulaXml = [...body.matchAll(formulaElementGlobalPattern)].find((formulaMatch) => {
      return readAttribute(formulaMatch[0], 't') === 'array'
    })?.[0]
    if (address && formulaXml) {
      formulas.push({ address, formulaXml })
    }
  }
  return formulas
}

function insertFormulaIntoCellBody(body: string, formulaXml: string): string {
  if (formulaElementPattern.test(body)) {
    return body.replace(formulaElementPattern, formulaXml)
  }
  const valueIndex = body.search(/<(?:v|is|extLst)\b/u)
  return valueIndex >= 0 ? `${body.slice(0, valueIndex)}${formulaXml}${body.slice(valueIndex)}` : `${formulaXml}${body}`
}

function upsertFormulaInExistingCell(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string | null {
  const addressPattern = escapeRegExp(formula.address)
  const cellPattern = new RegExp(`<c\\b(?<attributes>[^>]*\\br=(["'])${addressPattern}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/c>`, 'u')
  if (!cellPattern.test(sheetXml)) {
    return null
  }
  return sheetXml.replace(cellPattern, (_cellXml: string, attributes: string, _quote: string, body: string) => {
    return `<c${attributes}>${insertFormulaIntoCellBody(body, formula.formulaXml)}</c>`
  })
}

function insertFormulaCellIntoRow(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string | null {
  const rowNumber = cellRowNumber(formula.address)
  if (rowNumber === null) {
    return null
  }
  const rowPattern = new RegExp(`<row\\b(?<attributes>[^>]*\\br=(["'])${String(rowNumber)}\\2[^>]*)>(?<body>[\\s\\S]*?)<\\/row>`, 'u')
  if (!rowPattern.test(sheetXml)) {
    return null
  }
  const nextCellXml = `<c r="${escapeXml(formula.address)}">${formula.formulaXml}</c>`
  return sheetXml.replace(rowPattern, (_rowXml: string, attributes: string, _quote: string, body: string) => {
    const targetColumn = cellColumnIndex(formula.address)
    let insertIndex = body.length
    for (const match of body.matchAll(/<c\b(?<attributes>[^>]*)>(?:[\s\S]*?)<\/c>/gu)) {
      const address = readAttribute(match.groups?.['attributes'] ?? '', 'r')
      if (address && cellColumnIndex(address) > targetColumn) {
        insertIndex = match.index
        break
      }
    }
    return `<row${attributes}>${body.slice(0, insertIndex)}${nextCellXml}${body.slice(insertIndex)}</row>`
  })
}

function insertFormulaCellIntoSheetData(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string {
  const rowNumber = cellRowNumber(formula.address) ?? 1
  const rowXml = `<row r="${String(rowNumber)}"><c r="${escapeXml(formula.address)}">${formula.formulaXml}</c></row>`
  return sheetXml.includes('</sheetData>') ? sheetXml.replace('</sheetData>', `${rowXml}</sheetData>`) : sheetXml
}

function upsertArrayFormula(sheetXml: string, formula: WorkbookSheetArrayFormulaSnapshot): string {
  return (
    upsertFormulaInExistingCell(sheetXml, formula) ??
    insertFormulaCellIntoRow(sheetXml, formula) ??
    insertFormulaCellIntoSheetData(sheetXml, formula)
  )
}

export function readImportedWorkbookArrayFormulas(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetArrayFormulasSnapshot> {
  const zip = readXlsxZipEntries(source)
  const formulasBySheet = new Map<string, WorkbookSheetArrayFormulasSnapshot>()
  sheetNames.forEach((sheetName, sheetIndex) => {
    const formulas = readArrayFormulaSnapshots(getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`))
    if (formulas.length > 0) {
      formulasBySheet.set(sheetName, { formulas })
    }
  })
  return formulasBySheet
}

export function addExportArrayFormulasToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheetsWithArrayFormulas = snapshot.sheets.filter((sheet) => (sheet.metadata?.arrayFormulas?.formulas.length ?? 0) > 0)
  if (sheetsWithArrayFormulas.length === 0) {
    return bytes
  }

  const zip: XlsxZipEntries = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const arrayFormulas = sheet.metadata?.arrayFormulas?.formulas
      if (!arrayFormulas || arrayFormulas.length === 0) {
        return
      }
      const sheetPath = normalizeZipPath(`xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const nextSheetXml = arrayFormulas.reduce(upsertArrayFormula, sheetXml)
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
