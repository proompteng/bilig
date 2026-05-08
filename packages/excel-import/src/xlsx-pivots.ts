import { unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  LiteralInput,
  PivotAggregation,
  WorkbookDefinedNameSnapshot,
  WorkbookPivotArtifactsSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'
import {
  addContentTypeOverride,
  addExportPreservedPivotArtifactsToXlsxBytes,
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  escapeXml,
  nextRelationshipId,
  officeRelationshipNamespace,
  parseRelationships,
  pivotCacheDefinitionContentType,
  pivotCacheDefinitionRelationshipType,
  pivotCacheRecordsContentType,
  pivotCacheRecordsRelationshipType,
  pivotTableContentType,
  pivotTableRelationshipType,
  readImportedPivotArtifacts,
  resolveTargetPath,
  setZipText,
  spreadsheetNamespace,
} from './xlsx-pivot-artifacts.js'
import { getZipText, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = XlsxZipEntries

interface PivotCacheField {
  readonly name: string
  readonly values: readonly LiteralInput[]
}

interface PivotCacheTable {
  readonly fields: readonly PivotCacheField[]
  readonly rows: readonly (readonly LiteralInput[])[]
}

interface ParsedPivotCache {
  readonly cacheId: number
  readonly source: CellRangeRef
  readonly fields: readonly string[]
}

export interface ImportedWorkbookPivots {
  readonly pivots: WorkbookPivotSnapshot[] | undefined
  readonly hasExternalPivotCaches: boolean
  readonly artifacts: WorkbookPivotArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetPivotArtifactsSnapshot>
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

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

function numberAttribute(value: unknown): number | null {
  const number = typeof value === 'number' ? value : typeof value === 'string' ? Number(value) : Number.NaN
  return Number.isFinite(number) ? number : null
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

function absoluteAddress(address: string): string {
  const decoded = XLSX.utils.decode_cell(address)
  return `$${XLSX.utils.encode_col(decoded.c)}$${decoded.r + 1}`
}

function rangeRefA1(range: CellRangeRef): string {
  const start = absoluteAddress(range.startAddress).replaceAll('$', '')
  const end = absoluteAddress(range.endAddress).replaceAll('$', '')
  return start === end ? start : `${start}:${end}`
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

function pivotOutputRange(pivot: WorkbookPivotSnapshot): string {
  const start = XLSX.utils.decode_cell(pivot.address)
  const end = {
    r: start.r + Math.max(1, pivot.rows) - 1,
    c: start.c + Math.max(1, pivot.cols) - 1,
  }
  return XLSX.utils.encode_range({ s: start, e: end })
}

function expandWorksheetDimension(sheetXml: string, pivot: WorkbookPivotSnapshot): string {
  const match = /<dimension\b[^>]*\bref="([^"]+)"[^>]*\/>/u.exec(sheetXml)
  if (!match) {
    return sheetXml
  }
  try {
    const existing = XLSX.utils.decode_range(match[1] ?? 'A1')
    const next = XLSX.utils.decode_range(pivotOutputRange(pivot))
    const expanded = XLSX.utils.encode_range({
      s: { r: Math.min(existing.s.r, next.s.r), c: Math.min(existing.s.c, next.s.c) },
      e: { r: Math.max(existing.e.r, next.e.r), c: Math.max(existing.e.c, next.e.c) },
    })
    return sheetXml.replace(match[0], `<dimension ref="${expanded}"/>`)
  } catch {
    return sheetXml
  }
}

function addWorksheetPivotTableDefinition(sheetXml: string, relationshipId: string, pivot: WorkbookPivotSnapshot): string {
  const withNamespace = ensureRelationshipNamespace(expandWorksheetDimension(sheetXml, pivot))
  return withNamespace.replace('</worksheet>', `<pivotTableDefinition r:id="${escapeXml(relationshipId)}"/></worksheet>`)
}

function addWorkbookPivotCache(workbookXml: string, cacheId: number, relationshipId: string): string {
  const withNamespace = ensureRelationshipNamespace(workbookXml)
  const entry = `<pivotCache cacheId="${String(cacheId)}" r:id="${escapeXml(relationshipId)}"/>`
  if (/<pivotCaches\b/u.test(withNamespace)) {
    return withNamespace.replace('</pivotCaches>', `${entry}</pivotCaches>`)
  }
  const pivotCaches = `<pivotCaches>${entry}</pivotCaches>`
  if (withNamespace.includes('</definedNames>')) {
    return withNamespace.replace('</definedNames>', `</definedNames>${pivotCaches}`)
  }
  if (withNamespace.includes('</sheets>')) {
    return withNamespace.replace('</sheets>', `</sheets>${pivotCaches}`)
  }
  return withNamespace.replace('</workbook>', `${pivotCaches}</workbook>`)
}

function buildCellValueMap(sheet: WorkbookSnapshot['sheets'][number]): Map<string, LiteralInput> {
  const values = new Map<string, LiteralInput>()
  for (const cell of sheet.cells) {
    values.set(cell.address.toUpperCase(), cell.value ?? null)
  }
  return values
}

function readCellValue(cellsByAddress: ReadonlyMap<string, LiteralInput>, address: string): LiteralInput {
  return cellsByAddress.get(address.toUpperCase()) ?? null
}

function fallbackColumnName(index: number): string {
  return `Column ${String(index + 1)}`
}

function buildPivotCacheTable(snapshot: WorkbookSnapshot, pivot: WorkbookPivotSnapshot): PivotCacheTable | null {
  const sourceSheet = snapshot.sheets.find((sheet) => sheet.name === pivot.source.sheetName)
  if (!sourceSheet) {
    return null
  }
  let source: XLSX.Range
  try {
    source = XLSX.utils.decode_range(`${pivot.source.startAddress}:${pivot.source.endAddress}`)
  } catch {
    return null
  }
  const cellsByAddress = buildCellValueMap(sourceSheet)
  const fields: PivotCacheField[] = []
  for (let column = source.s.c; column <= source.e.c; column += 1) {
    const headerAddress = XLSX.utils.encode_cell({ r: source.s.r, c: column })
    const rawHeader = readCellValue(cellsByAddress, headerAddress)
    const name = typeof rawHeader === 'string' && rawHeader.trim().length > 0 ? rawHeader.trim() : fallbackColumnName(column - source.s.c)
    const values: LiteralInput[] = []
    for (let row = source.s.r + 1; row <= source.e.r; row += 1) {
      values.push(readCellValue(cellsByAddress, XLSX.utils.encode_cell({ r: row, c: column })))
    }
    fields.push({ name, values })
  }
  const rows: LiteralInput[][] = []
  for (let row = source.s.r + 1; row <= source.e.r; row += 1) {
    const values: LiteralInput[] = []
    for (let column = source.s.c; column <= source.e.c; column += 1) {
      values.push(readCellValue(cellsByAddress, XLSX.utils.encode_cell({ r: row, c: column })))
    }
    rows.push(values)
  }
  return { fields, rows }
}

function uniqueValues(values: readonly LiteralInput[]): LiteralInput[] {
  const seen = new Set<string>()
  const output: LiteralInput[] = []
  for (const value of values) {
    const key = `${typeof value}:${String(value)}`
    if (!seen.has(key)) {
      seen.add(key)
      output.push(value)
    }
  }
  return output
}

function cacheSharedItemXml(value: LiteralInput): string {
  if (value === null) {
    return '<m/>'
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<n v="${String(value)}"/>`
  }
  if (typeof value === 'boolean') {
    return `<b v="${value ? '1' : '0'}"/>`
  }
  return `<s v="${escapeXml(String(value ?? ''))}"/>`
}

function buildCacheFieldXml(field: PivotCacheField): string {
  const values = uniqueValues(field.values)
  return [
    `<cacheField name="${escapeXml(field.name)}" numFmtId="0">`,
    `<sharedItems count="${String(values.length)}">`,
    ...values.map(cacheSharedItemXml),
    '</sharedItems>',
    '</cacheField>',
  ].join('')
}

function buildPivotCacheDefinitionXml(
  pivot: WorkbookPivotSnapshot,
  cacheTable: PivotCacheTable,
  exportSourceSheetName: string,
  cacheRecordsRelationshipId: string,
): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotCacheDefinition xmlns="${spreadsheetNamespace}" xmlns:r="${officeRelationshipNamespace}" r:id="${escapeXml(
      cacheRecordsRelationshipId,
    )}" refreshOnLoad="1" refreshedVersion="8" createdVersion="8" minRefreshableVersion="3" recordCount="${String(cacheTable.rows.length)}">`,
    '<cacheSource type="worksheet">',
    `<worksheetSource ref="${escapeXml(rangeRefA1(pivot.source))}" sheet="${escapeXml(exportSourceSheetName)}"/>`,
    '</cacheSource>',
    `<cacheFields count="${String(cacheTable.fields.length)}">`,
    ...cacheTable.fields.map(buildCacheFieldXml),
    '</cacheFields>',
    '</pivotCacheDefinition>',
  ].join('')
}

function cacheRecordItemXml(value: LiteralInput): string {
  if (value === null) {
    return '<m/>'
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<n v="${String(value)}"/>`
  }
  if (typeof value === 'boolean') {
    return `<b v="${value ? '1' : '0'}"/>`
  }
  return `<s v="${escapeXml(String(value ?? ''))}"/>`
}

function buildPivotCacheRecordsXml(cacheTable: PivotCacheTable): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotCacheRecords xmlns="${spreadsheetNamespace}" count="${String(cacheTable.rows.length)}">`,
    ...cacheTable.rows.map((row) => ['<r>', ...row.map(cacheRecordItemXml), '</r>'].join('')),
    '</pivotCacheRecords>',
  ].join('')
}

function pivotFieldXml(index: number, fields: readonly string[], pivot: WorkbookPivotSnapshot): string {
  const fieldName = fields[index] ?? ''
  if (pivot.groupBy.includes(fieldName)) {
    return '<pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField>'
  }
  if (pivot.values.some((value) => value.sourceColumn === fieldName)) {
    return '<pivotField dataField="1" showAll="0"/>'
  }
  return '<pivotField showAll="0"/>'
}

function buildRowFieldsXml(fieldIndexes: readonly number[]): string {
  if (fieldIndexes.length === 0) {
    return ''
  }
  return [
    `<rowFields count="${String(fieldIndexes.length)}">`,
    ...fieldIndexes.map((index) => `<field x="${String(index)}"/>`),
    '</rowFields>',
  ].join('')
}

function subtotalValue(value: PivotAggregation): string {
  switch (value) {
    case 'sum':
      return 'sum'
    case 'count':
      return 'count'
  }
}

function defaultDataFieldName(value: WorkbookPivotValueSnapshot): string {
  if (value.outputLabel && value.outputLabel.trim().length > 0) {
    return value.outputLabel.trim()
  }
  return `${value.summarizeBy === 'sum' ? 'Sum' : 'Count'} of ${value.sourceColumn}`
}

function buildDataFieldsXml(values: readonly WorkbookPivotValueSnapshot[], fieldIndexesByName: ReadonlyMap<string, number>): string {
  if (values.length === 0) {
    return ''
  }
  const fields = values.flatMap((value) => {
    const fieldIndex = fieldIndexesByName.get(value.sourceColumn)
    if (fieldIndex === undefined) {
      return []
    }
    return [
      `<dataField name="${escapeXml(defaultDataFieldName(value))}" fld="${String(fieldIndex)}" subtotal="${subtotalValue(
        value.summarizeBy,
      )}"/>`,
    ]
  })
  return fields.length > 0 ? [`<dataFields count="${String(fields.length)}">`, ...fields, '</dataFields>'].join('') : ''
}

function buildPivotTableDefinitionXml(pivot: WorkbookPivotSnapshot, cacheId: number, cacheTable: PivotCacheTable): string {
  const fieldNames = cacheTable.fields.map((field) => field.name)
  const fieldIndexesByName = new Map(fieldNames.map((name, index) => [name, index]))
  const rowFieldIndexes = pivot.groupBy.flatMap((fieldName) => {
    const index = fieldIndexesByName.get(fieldName)
    return index === undefined ? [] : [index]
  })
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<pivotTableDefinition xmlns="${spreadsheetNamespace}" name="${escapeXml(pivot.name)}" cacheId="${String(
      cacheId,
    )}" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">`,
    `<location ref="${escapeXml(pivotOutputRange(pivot))}" firstHeaderRow="1" firstDataRow="2" firstDataCol="${String(
      Math.max(1, rowFieldIndexes.length),
    )}"/>`,
    `<pivotFields count="${String(fieldNames.length)}">`,
    ...fieldNames.map((_, index) => pivotFieldXml(index, fieldNames, pivot)),
    '</pivotFields>',
    buildRowFieldsXml(rowFieldIndexes),
    buildDataFieldsXml(pivot.values, fieldIndexesByName),
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>',
    '</pivotTableDefinition>',
  ].join('')
}

export function addExportPivotsToXlsxBytes(
  bytes: Uint8Array,
  snapshot: WorkbookSnapshot,
  exportSheetNamesByOriginalName: ReadonlyMap<string, string>,
): Uint8Array {
  if (snapshot.workbook.metadata?.pivotArtifacts) {
    return addExportPreservedPivotArtifactsToXlsxBytes(bytes, snapshot)
  }
  const pivots = snapshot.workbook.metadata?.pivots ?? []
  if (pivots.length === 0) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let nextPivotTableIndex = nextPartIndex(zip, 'xl/pivotTables/pivotTable', '.xml')
  let nextPivotCacheIndex = nextPartIndex(zip, 'xl/pivotCache/pivotCacheDefinition', '.xml')
  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  let workbookXml = getZipText(zip, 'xl/workbook.xml') ?? ''
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetPivots = pivots.filter((pivot) => pivot.sheetName === sheet.name)
      if (sheetPivots.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      let sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      let updatedSheetXml = sheetXml
      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))

      sheetPivots.forEach((pivot) => {
        const cacheTable = buildPivotCacheTable(snapshot, pivot)
        const exportSourceSheetName = exportSheetNamesByOriginalName.get(pivot.source.sheetName) ?? pivot.source.sheetName
        if (!cacheTable || cacheTable.fields.length === 0 || pivot.values.length === 0 || !exportSourceSheetName) {
          return
        }
        const pivotTableIndex = nextPivotTableIndex
        nextPivotTableIndex += 1
        const cacheIndex = nextPivotCacheIndex
        nextPivotCacheIndex += 1
        const pivotTablePath = `xl/pivotTables/pivotTable${String(pivotTableIndex)}.xml`
        const cacheDefinitionPath = `xl/pivotCache/pivotCacheDefinition${String(cacheIndex)}.xml`
        const cacheRecordsPath = `xl/pivotCache/pivotCacheRecords${String(cacheIndex)}.xml`
        const cacheRecordsRelationshipId = 'rId1'

        const workbookRelationshipId = nextRelationshipId(workbookRelationships)
        workbookRelationships.push({
          id: workbookRelationshipId,
          type: pivotCacheDefinitionRelationshipType,
          target: `pivotCache/pivotCacheDefinition${String(cacheIndex)}.xml`,
        })
        workbookXml = addWorkbookPivotCache(workbookXml, cacheIndex, workbookRelationshipId)

        setZipText(zip, pivotTablePath, buildPivotTableDefinitionXml(pivot, cacheIndex, cacheTable))
        setZipText(
          zip,
          cacheDefinitionPath,
          buildPivotCacheDefinitionXml(pivot, cacheTable, exportSourceSheetName, cacheRecordsRelationshipId),
        )
        setZipText(zip, cacheRecordsPath, buildPivotCacheRecordsXml(cacheTable))
        setZipText(
          zip,
          `xl/pivotCache/_rels/pivotCacheDefinition${String(cacheIndex)}.xml.rels`,
          buildRelationshipsXml([
            {
              id: cacheRecordsRelationshipId,
              type: pivotCacheRecordsRelationshipType,
              target: `pivotCacheRecords${String(cacheIndex)}.xml`,
            },
          ]),
        )

        const sheetRelationshipId = nextRelationshipId(sheetRelationships)
        sheetRelationships.push({
          id: sheetRelationshipId,
          type: pivotTableRelationshipType,
          target: `../pivotTables/pivotTable${String(pivotTableIndex)}.xml`,
        })
        updatedSheetXml = addWorksheetPivotTableDefinition(updatedSheetXml, sheetRelationshipId, pivot)

        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${pivotTablePath}`, pivotTableContentType)
        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${cacheDefinitionPath}`, pivotCacheDefinitionContentType)
        contentTypesXml = addContentTypeOverride(contentTypesXml, `/${cacheRecordsPath}`, pivotCacheRecordsContentType)
      })

      setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      setZipText(zip, sheetPath, updatedSheetXml)
    })

  if (workbookXml.length > 0) {
    setZipText(zip, 'xl/workbook.xml', workbookXml)
  }
  setZipText(zip, 'xl/_rels/workbook.xml.rels', buildRelationshipsXml(workbookRelationships))
  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }
  return zipSync(zip)
}

function readWorkbookPivotCaches(zip: ZipEntries): Map<number, string> {
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const output = new Map<number, string>()
  if (!workbookXml) {
    return output
  }
  const parsed: unknown = xmlParser.parse(workbookXml)
  const pivotCaches = asArray(recordChild(recordChild(parsed, 'workbook'), 'pivotCaches')?.['pivotCache'])
  pivotCaches.forEach((entry) => {
    if (!isRecord(entry)) {
      return
    }
    const cacheId = numberAttribute(entry['cacheId'])
    const relationshipId = typeof entry['id'] === 'string' ? entry['id'] : null
    const relationship = relationshipId
      ? workbookRelationships.find(
          (candidate) => candidate.id === relationshipId && candidate.type === pivotCacheDefinitionRelationshipType,
        )
      : undefined
    if (cacheId !== null && relationship) {
      output.set(cacheId, resolveTargetPath('xl/workbook.xml', relationship.target))
    }
  })
  return output
}

function parsePivotCacheDefinition(
  cacheId: number,
  xml: string,
  tablesByName: ReadonlyMap<string, WorkbookTableSnapshot>,
  definedNamesByName: ReadonlyMap<string, WorkbookDefinedNameSnapshot>,
): ParsedPivotCache | null {
  const parsed: unknown = xmlParser.parse(xml)
  const definition = recordChild(parsed, 'pivotCacheDefinition')
  const sourceRecord = recordChild(recordChild(definition, 'cacheSource'), 'worksheetSource')
  const sheetName = typeof sourceRecord?.['sheet'] === 'string' ? sourceRecord['sheet'] : null
  const ref = typeof sourceRecord?.['ref'] === 'string' ? sourceRecord['ref'] : null
  const sourceName = typeof sourceRecord?.['name'] === 'string' ? sourceRecord['name'].trim() : ''
  const namedSource = sourceName.length > 0 ? sourceRangeForName(sourceName, sheetName, tablesByName, definedNamesByName) : null
  const source = sheetName && ref ? parseRangeRef(sheetName, ref) : namedSource
  if (!source) {
    return null
  }
  const fields = asArray(recordChild(definition, 'cacheFields')?.['cacheField']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || entry['name'].trim().length === 0) {
      return []
    }
    return [entry['name'].trim()]
  })
  return fields.length > 0 ? { cacheId, source, fields } : null
}

function pivotCacheDefinitionHasExternalSource(xml: string): boolean {
  const parsed: unknown = xmlParser.parse(xml)
  const cacheSource = recordChild(recordChild(parsed, 'pivotCacheDefinition'), 'cacheSource')
  return cacheSource?.['type'] === 'external'
}

function sourceRangeForName(
  name: string,
  sheetName: string | null,
  tablesByName: ReadonlyMap<string, WorkbookTableSnapshot>,
  definedNamesByName: ReadonlyMap<string, WorkbookDefinedNameSnapshot>,
): CellRangeRef | null {
  const normalizedName = name.toLocaleLowerCase('en-US')
  const table = tablesByName.get(normalizedName)
  if (table) {
    return {
      sheetName: table.sheetName,
      startAddress: table.startAddress,
      endAddress: table.endAddress,
    }
  }
  const definedName = sheetName
    ? (definedNamesByName.get(definedNameKey(name, sheetName)) ?? definedNamesByName.get(normalizedName))
    : definedNamesByName.get(normalizedName)
  if (!definedName) {
    return null
  }
  const value = definedName.value
  if (!isRecord(value)) {
    return null
  }
  if (
    value['kind'] === 'range-ref' &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  ) {
    return { sheetName: value['sheetName'], startAddress: value['startAddress'], endAddress: value['endAddress'] }
  }
  if (value['kind'] === 'cell-ref' && typeof value['sheetName'] === 'string' && typeof value['address'] === 'string') {
    return { sheetName: value['sheetName'], startAddress: value['address'], endAddress: value['address'] }
  }
  return null
}

function definedNameKey(name: string, scopeSheetName: string | undefined): string {
  return scopeSheetName
    ? `${scopeSheetName.toLocaleLowerCase('en-US')}:${name.toLocaleLowerCase('en-US')}`
    : name.toLocaleLowerCase('en-US')
}

function parsePivotCaches(
  zip: ZipEntries,
  tables: readonly WorkbookTableSnapshot[],
  definedNames: readonly WorkbookDefinedNameSnapshot[],
): { readonly caches: Map<number, ParsedPivotCache>; readonly hasExternalPivotCaches: boolean } {
  const cacheDefinitions = readWorkbookPivotCaches(zip)
  const tablesByName = new Map(tables.map((table) => [table.name.toLocaleLowerCase('en-US'), table]))
  const definedNamesByName = new Map(
    definedNames.map((definedName) => [definedNameKey(definedName.name, definedName.scopeSheetName), definedName]),
  )
  const output = new Map<number, ParsedPivotCache>()
  let hasExternalPivotCaches = false
  for (const [cacheId, path] of cacheDefinitions.entries()) {
    const xml = getZipText(zip, path) ?? ''
    hasExternalPivotCaches ||= pivotCacheDefinitionHasExternalSource(xml)
    const parsed = parsePivotCacheDefinition(cacheId, xml, tablesByName, definedNamesByName)
    if (parsed) {
      output.set(cacheId, parsed)
    }
  }
  return { caches: output, hasExternalPivotCaches }
}

function aggregationFromSubtotal(value: unknown): PivotAggregation | null {
  switch (value) {
    case undefined:
    case null:
    case 'sum':
      return 'sum'
    case 'count':
    case 'countNums':
      return 'count'
    default:
      return null
  }
}

function parsePivotTableXml(sheetName: string, xml: string, caches: ReadonlyMap<number, ParsedPivotCache>): WorkbookPivotSnapshot | null {
  const parsed: unknown = xmlParser.parse(xml)
  const definition = recordChild(parsed, 'pivotTableDefinition')
  const cacheId = numberAttribute(definition?.['cacheId'])
  const cache = cacheId === null ? undefined : caches.get(cacheId)
  const locationRefValue = recordChild(definition, 'location')?.['ref']
  const locationRef = typeof locationRefValue === 'string' ? locationRefValue : null
  if (!definition || !cache || !locationRef) {
    return null
  }
  const location = parseRangeRef(sheetName, locationRef)
  if (!location) {
    return null
  }
  const groupBy = asArray(recordChild(definition, 'rowFields')?.['field']).flatMap((entry) => {
    const index = isRecord(entry) ? numberAttribute(entry['x']) : null
    const field = index === null ? undefined : cache.fields[index]
    return field ? [field] : []
  })
  const values = asArray(recordChild(definition, 'dataFields')?.['dataField']).flatMap((entry) => {
    if (!isRecord(entry)) {
      return []
    }
    const fieldIndex = numberAttribute(entry['fld'])
    const sourceColumn = fieldIndex === null ? undefined : cache.fields[fieldIndex]
    const summarizeBy = aggregationFromSubtotal(entry['subtotal'])
    if (!sourceColumn || !summarizeBy) {
      return []
    }
    const outputLabel = typeof entry['name'] === 'string' && entry['name'].trim().length > 0 ? entry['name'].trim() : undefined
    const value: WorkbookPivotValueSnapshot = { sourceColumn, summarizeBy }
    if (outputLabel !== undefined) {
      value.outputLabel = outputLabel
    }
    return [value]
  })
  if (values.length === 0) {
    return null
  }
  const name =
    typeof definition['name'] === 'string' && definition['name'].trim().length > 0
      ? definition['name'].trim()
      : `Pivot ${location.startAddress}`
  const start = XLSX.utils.decode_cell(location.startAddress)
  const end = XLSX.utils.decode_cell(location.endAddress)
  return {
    name,
    sheetName,
    address: location.startAddress,
    source: cache.source,
    groupBy,
    values,
    rows: Math.max(1, end.r - start.r + 1),
    cols: Math.max(1, end.c - start.c + 1),
  }
}

export function readImportedWorkbookPivots(
  source: XlsxZipSource,
  sheetNames: readonly string[],
  tables: readonly WorkbookTableSnapshot[] = [],
  definedNames: readonly WorkbookDefinedNameSnapshot[] = [],
): ImportedWorkbookPivots {
  const zip = readXlsxZipEntries(source)
  const { artifacts, sheetArtifactsByName } = readImportedPivotArtifacts(zip, sheetNames)
  const { caches, hasExternalPivotCaches } = parsePivotCaches(zip, tables, definedNames)
  if (caches.size === 0) {
    return { pivots: undefined, hasExternalPivotCaches, artifacts, sheetArtifactsByName }
  }
  const pivots: WorkbookPivotSnapshot[] = []
  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml || !/<pivotTableDefinition\b/u.test(sheetXml)) {
      return
    }
    const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
    const parsedSheet: unknown = xmlParser.parse(sheetXml)
    const pivotRefs = asArray(recordChild(parsedSheet, 'worksheet')?.['pivotTableDefinition'])
    pivotRefs.forEach((entry) => {
      if (!isRecord(entry) || typeof entry['id'] !== 'string') {
        return
      }
      const relationship = sheetRelationships.find(
        (candidate) => candidate.id === entry['id'] && candidate.type === pivotTableRelationshipType,
      )
      if (!relationship) {
        return
      }
      const pivotPath = resolveTargetPath(sheetPath, relationship.target)
      const pivot = parsePivotTableXml(sheetName, getZipText(zip, pivotPath) ?? '', caches)
      if (pivot) {
        pivots.push(pivot)
      }
    })
  })
  return {
    pivots:
      pivots.length > 0
        ? pivots.toSorted((left, right) =>
            `${left.sheetName}:${left.address}:${left.name}`.localeCompare(`${right.sheetName}:${right.address}:${right.name}`),
          )
        : undefined,
    hasExternalPivotCaches,
    artifacts,
    sheetArtifactsByName,
  }
}
