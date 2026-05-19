import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot, WorkbookTableColumnSnapshot, WorkbookTableSnapshot, WorkbookTableStyleSnapshot } from '@bilig/protocol'
import { decodeExcelEscapedText, encodeExcelEscapedText } from './xlsx-escaped-text.js'
import { readSortStateXml } from './xlsx-sorts.js'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const tableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
const tableContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'

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

function escapeExcelXmlTextAttribute(value: string): string {
  return escapeXml(encodeExcelEscapedText(value))
}

function parseBooleanAttribute(value: unknown): boolean | undefined {
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  return undefined
}

function booleanAttribute(value: boolean): string {
  return value ? '1' : '0'
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

function nextPartIndex(zip: ZipEntries, prefix: string, suffix: string): number {
  let next = 1
  for (const path of Object.keys(zip)) {
    if (!path.startsWith(prefix) || !path.endsWith(suffix)) {
      continue
    }
    const raw = path.slice(prefix.length, -suffix.length)
    const value = Number(raw)
    if (Number.isInteger(value) && value >= next) {
      next = value + 1
    }
  }
  return next
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['Id'] !== 'string' || typeof entry['Target'] !== 'string' || typeof entry['Type'] !== 'string') {
      return []
    }
    return [{ id: entry['Id'], target: entry['Target'], type: entry['Type'] }]
  })
}

function nextRelationshipId(relationships: readonly ParsedRelationship[]): string {
  let next = 1
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/u.exec(relationship.id)
    if (match) {
      next = Math.max(next, Number(match[1]) + 1)
    }
  }
  return `rId${String(next)}`
}

function buildRelationshipsXml(relationships: readonly ParsedRelationship[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) =>
        `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(
          relationship.target,
        )}"/>`,
    ),
    '</Relationships>',
  ].join('')
}

function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}

function resolveTargetPath(basePartPath: string, target: string): string {
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}

function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<([A-Za-z0-9:]+)\b([^>]*)>/u, `<$1$2 xmlns:r="${officeRelationshipNamespace}">`)
}

function rangeRefA1(table: WorkbookTableSnapshot): string {
  const start = XLSX.utils.decode_cell(table.startAddress)
  const end = XLSX.utils.decode_cell(table.endAddress)
  return XLSX.utils.encode_range({ s: start, e: end })
}

function parseRangeRef(ref: string): { startAddress: string; endAddress: string } | null {
  try {
    const decoded = XLSX.utils.decode_range(ref.replaceAll('$', ''))
    return {
      startAddress: XLSX.utils.encode_cell(decoded.s),
      endAddress: XLSX.utils.encode_cell(decoded.e),
    }
  } catch {
    return null
  }
}

function addWorksheetTablePart(sheetXml: string, relationshipId: string): string {
  const withNamespace = ensureRelationshipNamespace(sheetXml)
  const tablePart = `<tablePart r:id="${escapeXml(relationshipId)}"/>`
  const tablePartsMatch = /<tableParts\b[^>]*\bcount="(\d+)"[^>]*>([\s\S]*?)<\/tableParts>/u.exec(withNamespace)
  if (tablePartsMatch) {
    const nextCount = Number(tablePartsMatch[1] ?? '0') + 1
    return withNamespace.replace(
      tablePartsMatch[0],
      `<tableParts count="${String(nextCount)}">${tablePartsMatch[2] ?? ''}${tablePart}</tableParts>`,
    )
  }
  return withNamespace.replace('</worksheet>', `<tableParts count="1">${tablePart}</tableParts></worksheet>`)
}

function tableDisplayName(name: string, fallbackIndex: number): string {
  const sanitized = name
    .trim()
    .replace(/[^\p{L}\p{N}_.]/gu, '_')
    .replace(/^[^\p{L}_]+/u, '')
  return sanitized.length > 0 ? sanitized : `Table${String(fallbackIndex)}`
}

function exportTableColumns(table: WorkbookTableSnapshot): WorkbookTableColumnSnapshot[] {
  const columnNames = table.columnNames.length > 0 ? table.columnNames : ['Column 1']
  return columnNames.map((columnName, index) => {
    const column = table.columns?.[index]
    const exportedColumn: WorkbookTableColumnSnapshot = {
      name: typeof column?.name === 'string' && column.name.trim().length > 0 ? column.name : columnName,
    }
    if (column?.totalsRowLabel !== undefined) {
      exportedColumn.totalsRowLabel = column.totalsRowLabel
    }
    if (column?.totalsRowFunction !== undefined) {
      exportedColumn.totalsRowFunction = column.totalsRowFunction
    }
    return exportedColumn
  })
}

function buildTableColumnXml(column: WorkbookTableColumnSnapshot, index: number): string {
  return [
    `<tableColumn id="${String(index + 1)}" name="${escapeExcelXmlTextAttribute(column.name)}"`,
    column.totalsRowLabel !== undefined ? ` totalsRowLabel="${escapeExcelXmlTextAttribute(column.totalsRowLabel)}"` : '',
    column.totalsRowFunction !== undefined ? ` totalsRowFunction="${escapeXml(column.totalsRowFunction)}"` : '',
    '/>',
  ].join('')
}

function buildTableStyleXml(style: WorkbookTableStyleSnapshot | undefined): string {
  const name = style?.name ?? 'TableStyleMedium2'
  const showFirstColumn = style?.showFirstColumn ?? false
  const showLastColumn = style?.showLastColumn ?? false
  const showRowStripes = style?.showRowStripes ?? true
  const showColumnStripes = style?.showColumnStripes ?? false
  return [
    `<tableStyleInfo name="${escapeXml(name)}"`,
    ` showFirstColumn="${booleanAttribute(showFirstColumn)}"`,
    ` showLastColumn="${booleanAttribute(showLastColumn)}"`,
    ` showRowStripes="${booleanAttribute(showRowStripes)}"`,
    ` showColumnStripes="${booleanAttribute(showColumnStripes)}"/>`,
  ].join('')
}

function buildTableXml(table: WorkbookTableSnapshot, tableId: number): string {
  const ref = rangeRefA1(table)
  const displayName = tableDisplayName(table.name, tableId)
  const columns = exportTableColumns(table)
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<table xmlns="${spreadsheetNamespace}" id="${String(tableId)}" name="${escapeXml(displayName)}" displayName="${escapeXml(
      displayName,
    )}" ref="${escapeXml(ref)}" headerRowCount="${table.headerRow ? '1' : '0'}" totalsRowShown="${table.totalsRow ? '1' : '0'}">`,
    table.headerRow ? `<autoFilter ref="${escapeXml(ref)}"/>` : '',
    table.sortState ?? '',
    `<tableColumns count="${String(columns.length)}">`,
    ...columns.map(buildTableColumnXml),
    '</tableColumns>',
    buildTableStyleXml(table.style),
    '</table>',
  ].join('')
}

export function addExportTablesToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): Uint8Array {
  const tables = snapshot.workbook.metadata?.tables ?? []
  if (tables.length === 0) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let nextTableIndex = nextPartIndex(zip, 'xl/tables/table', '.xml')
  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetTables = tables.filter((table) => table.sheetName === sheet.name)
      if (sheetTables.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml || !exportSheetNamesByOriginalName.has(sheet.name)) {
        return
      }
      let updatedSheetXml = sheetXml
      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))

      sheetTables.forEach((table) => {
        const tableIndex = nextTableIndex
        nextTableIndex += 1
        const tablePath = `xl/tables/table${String(tableIndex)}.xml`
        setZipText(zip, tablePath, buildTableXml(table, tableIndex))
        const relationshipId = nextRelationshipId(sheetRelationships)
        sheetRelationships.push({
          id: relationshipId,
          type: tableRelationshipType,
          target: `../tables/table${String(tableIndex)}.xml`,
        })
        updatedSheetXml = addWorksheetTablePart(updatedSheetXml, relationshipId)
        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${tablePath}`, tableContentType)
      })

      setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      setZipText(zip, sheetPath, updatedSheetXml)
    })

  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }
  return zipSync(zip)
}

function parseTableXml(sheetName: string, xml: string): WorkbookTableSnapshot | null {
  const parsed: unknown = xmlParser.parse(xml)
  const table = recordChild(parsed, 'table')
  const ref = typeof table?.['ref'] === 'string' ? parseRangeRef(table['ref']) : null
  if (!table || !ref) {
    return null
  }
  const name =
    typeof table['displayName'] === 'string' && table['displayName'].trim().length > 0
      ? table['displayName'].trim()
      : typeof table['name'] === 'string' && table['name'].trim().length > 0
        ? table['name'].trim()
        : `Table_${ref.startAddress}`
  const columns = asArray(recordChild(table, 'tableColumns')?.['tableColumn']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string') {
      return []
    }
    const columnName = decodeExcelEscapedText(entry['name']).trim()
    if (columnName.length === 0) {
      return []
    }
    const column: WorkbookTableColumnSnapshot = {
      name: columnName,
      ...(typeof entry['totalsRowLabel'] === 'string' ? { totalsRowLabel: decodeExcelEscapedText(entry['totalsRowLabel']) } : {}),
      ...(typeof entry['totalsRowFunction'] === 'string' ? { totalsRowFunction: entry['totalsRowFunction'] } : {}),
    }
    return [column]
  })
  const hasTotalsRowFormula = asArray(recordChild(table, 'tableColumns')?.['tableColumn']).some(
    (entry) => isRecord(entry) && (typeof entry['totalsRowFormula'] === 'string' || isRecord(entry['totalsRowFormula'])),
  )
  const style = parseTableStyle(recordChild(table, 'tableStyleInfo'))
  const sortState = readSortStateXml(xml)
  const explicitTotalsRow = parseBooleanAttribute(table['totalsRowShown']) ?? parseBooleanAttribute(table['totalsRowCount'])
  return {
    name,
    sheetName,
    startAddress: ref.startAddress,
    endAddress: ref.endAddress,
    columnNames: columns.map((column) => column.name),
    ...(columns.some((column) => column.totalsRowLabel !== undefined || column.totalsRowFunction !== undefined) ? { columns } : {}),
    headerRow: table['headerRowCount'] !== '0',
    totalsRow: explicitTotalsRow ?? hasTotalsRowFormula,
    ...(style ? { style } : {}),
    ...(sortState ? { sortState } : {}),
  }
}

function parseTableStyle(style: Record<string, unknown> | null): WorkbookTableStyleSnapshot | undefined {
  if (!style) {
    return undefined
  }
  const parsed: WorkbookTableStyleSnapshot = {}
  if (typeof style['name'] === 'string' && style['name'].trim().length > 0) {
    parsed.name = style['name'].trim()
  }
  const showFirstColumn = parseBooleanAttribute(style['showFirstColumn'])
  const showLastColumn = parseBooleanAttribute(style['showLastColumn'])
  const showRowStripes = parseBooleanAttribute(style['showRowStripes'])
  const showColumnStripes = parseBooleanAttribute(style['showColumnStripes'])
  if (showFirstColumn !== undefined) {
    parsed.showFirstColumn = showFirstColumn
  }
  if (showLastColumn !== undefined) {
    parsed.showLastColumn = showLastColumn
  }
  if (showRowStripes !== undefined) {
    parsed.showRowStripes = showRowStripes
  }
  if (showColumnStripes !== undefined) {
    parsed.showColumnStripes = showColumnStripes
  }
  if (Object.keys(parsed).length === 0 || isDefaultTableStyle(parsed)) {
    return undefined
  }
  return parsed
}

function isDefaultTableStyle(style: WorkbookTableStyleSnapshot): boolean {
  return (
    (style.name ?? 'TableStyleMedium2') === 'TableStyleMedium2' &&
    !(style.showFirstColumn ?? false) &&
    !(style.showLastColumn ?? false) &&
    (style.showRowStripes ?? true) &&
    !(style.showColumnStripes ?? false)
  )
}

export function readImportedWorkbookTables(source: XlsxZipSource, sheetNames: readonly string[]): WorkbookTableSnapshot[] | undefined {
  const zip = readXlsxZipEntries(source)
  const tables: WorkbookTableSnapshot[] = []

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml || !/<tableParts\b/u.test(sheetXml)) {
      return
    }
    const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
    const parsedSheet: unknown = xmlParser.parse(sheetXml)
    const tableRefs = asArray(recordChild(recordChild(parsedSheet, 'worksheet'), 'tableParts')?.['tablePart'])
    tableRefs.forEach((entry) => {
      if (!isRecord(entry) || typeof entry['id'] !== 'string') {
        return
      }
      const relationship = sheetRelationships.find((candidate) => candidate.id === entry['id'] && candidate.type === tableRelationshipType)
      if (!relationship) {
        return
      }
      const tablePath = resolveTargetPath(sheetPath, relationship.target)
      const table = parseTableXml(sheetName, getZipText(zip, tablePath) ?? '')
      if (table) {
        tables.push(table)
      }
    })
  })

  return tables.length > 0 ? tables.toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
}
