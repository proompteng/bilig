import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot, WorkbookTableSnapshot } from '@bilig/protocol'
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
    .replace(/[^A-Za-z0-9_.]/g, '_')
    .replace(/^[^A-Za-z_]+/u, '')
  return sanitized.length > 0 ? sanitized : `Table${String(fallbackIndex)}`
}

function buildTableXml(table: WorkbookTableSnapshot, tableId: number): string {
  const ref = rangeRefA1(table)
  const displayName = tableDisplayName(table.name, tableId)
  const columns = table.columnNames.length > 0 ? table.columnNames : ['Column 1']
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<table xmlns="${spreadsheetNamespace}" id="${String(tableId)}" name="${escapeXml(displayName)}" displayName="${escapeXml(
      displayName,
    )}" ref="${escapeXml(ref)}" headerRowCount="${table.headerRow ? '1' : '0'}" totalsRowShown="${table.totalsRow ? '1' : '0'}">`,
    table.headerRow ? `<autoFilter ref="${escapeXml(ref)}"/>` : '',
    `<tableColumns count="${String(columns.length)}">`,
    ...columns.map((columnName, index) => `<tableColumn id="${String(index + 1)}" name="${escapeXml(columnName)}"/>`),
    '</tableColumns>',
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
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
  const columnNames = asArray(recordChild(table, 'tableColumns')?.['tableColumn']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || entry['name'].trim().length === 0) {
      return []
    }
    return [entry['name'].trim()]
  })
  return {
    name,
    sheetName,
    startAddress: ref.startAddress,
    endAddress: ref.endAddress,
    columnNames,
    headerRow: table['headerRowCount'] !== '0',
    totalsRow: table['totalsRowShown'] === '1' || table['totalsRowCount'] === '1',
  }
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
