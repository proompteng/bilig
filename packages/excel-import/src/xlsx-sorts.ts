import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { CellRangeRef, WorkbookSnapshot, WorkbookSortSnapshot } from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetSortStateTailElements = [
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'drawing',
  'legacyDrawing',
  'tableParts',
  'pivotTableDefinition',
] as const

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function asArray(value: unknown): unknown[] {
  if (value === undefined || value === null) {
    return []
  }
  return Array.isArray(value) ? value : [value]
}

function recordChild(value: unknown, key: string): Record<string, unknown> | null {
  if (!isRecord(value)) {
    return null
  }
  const child = value[key]
  return isRecord(child) ? child : null
}

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function getZipText(zip: ZipEntries, path: string): string | null {
  const file = zip[normalizeZipPath(path)]
  return file ? strFromU8(file) : null
}

function setZipText(zip: ZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function rangeRefA1(range: CellRangeRef): string | null {
  try {
    const decoded = XLSX.utils.decode_range(`${range.startAddress}:${range.endAddress}`.replaceAll('$', ''))
    return XLSX.utils.encode_range(decoded)
  } catch {
    return null
  }
}

function parseRangeRef(sheetName: string, ref: string): CellRangeRef | null {
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      sheetName,
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function isFalseAttribute(value: unknown): boolean {
  return value === false || value === '0' || value === 'false'
}

function keyConditionRef(range: CellRangeRef, keyAddress: string): string | null {
  try {
    const rangeBounds = XLSX.utils.decode_range(`${range.startAddress}:${range.endAddress}`.replaceAll('$', ''))
    const key = XLSX.utils.decode_cell(keyAddress.replaceAll('$', ''))
    if (key.c < rangeBounds.s.c || key.c > rangeBounds.e.c || key.r > rangeBounds.e.r) {
      return null
    }
    const startRow = key.r
    return XLSX.utils.encode_range({
      s: { r: startRow, c: key.c },
      e: { r: rangeBounds.e.r, c: key.c },
    })
  } catch {
    return null
  }
}

function exportSortConditionXml(sort: WorkbookSortSnapshot, key: WorkbookSortSnapshot['keys'][number]): string | null {
  const ref = keyConditionRef(sort.range, key.keyAddress)
  if (!ref) {
    return null
  }
  const descending = key.direction === 'desc' ? ' descending="1"' : ''
  return `<sortCondition${descending} ref="${escapeXml(ref)}"/>`
}

function insertWorksheetSortState(sheetXml: string, sortXml: string): string {
  if (/<sortState\b/u.test(sheetXml)) {
    return sheetXml.replace(/<sortState\b[^>]*(?:\/>|>[\s\S]*?<\/sortState>)/u, sortXml)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetSortStateTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${sortXml}${sheetXml.slice(insertIndex)}`
}

function exportSortStateXml(sheetName: string, sort: WorkbookSortSnapshot): string | null {
  const ref = sort.range.sheetName === sheetName ? rangeRefA1(sort.range) : null
  if (!ref) {
    return null
  }
  const conditions = sort.keys.flatMap((key) => {
    const xml = exportSortConditionXml(sort, key)
    return xml ? [xml] : []
  })
  if (conditions.length === 0) {
    return null
  }
  return `<sortState ref="${escapeXml(ref)}">${conditions.join('')}</sortState>`
}

export function addExportSortsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.sorts ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sort = (sheet.metadata?.sorts ?? []).find((candidate) => candidate.range.sheetName === sheet.name)
      const sortXml = sort ? exportSortStateXml(sheet.name, sort) : null
      if (!sortXml) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetSortState(sheetXml, sortXml))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

function parseSortKey(ref: unknown, direction: unknown): WorkbookSortSnapshot['keys'][number] | null {
  if (typeof ref !== 'string') {
    return null
  }
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      keyAddress: XLSX.utils.encode_cell(decoded.s),
      direction: isFalseAttribute(direction) || direction === undefined ? 'asc' : 'desc',
    }
  } catch {
    return null
  }
}

function parseSortState(sheetName: string, sortState: unknown): WorkbookSortSnapshot | null {
  if (!isRecord(sortState) || typeof sortState['ref'] !== 'string') {
    return null
  }
  const range = parseRangeRef(sheetName, sortState['ref'])
  if (!range) {
    return null
  }
  const keys = asArray(sortState['sortCondition']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const key = parseSortKey(entry['ref'], entry['descending'])
    return key ? [key] : []
  })
  return keys.length > 0 ? { range, keys } : null
}

export function readImportedWorkbookSorts(source: XlsxZipSource, sheetNames: readonly string[]): Map<string, WorkbookSortSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const sortsBySheet = new Map<string, WorkbookSortSnapshot[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<sortState\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const sorts = asArray(recordChild(recordChild(parsed, 'worksheet'), 'sortState')).flatMap((entry) => {
      const sort = parseSortState(sheetName, entry)
      return sort ? [sort] : []
    })
    if (sorts.length > 0) {
      sortsBySheet.set(sheetName, sorts)
    }
  })

  return sortsBySheet
}
