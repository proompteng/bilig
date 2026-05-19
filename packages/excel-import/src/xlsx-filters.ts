import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type {
  CellRangeRef,
  WorkbookAutoFilterColumnSnapshot,
  WorkbookAutoFilterCustomCriteriaSnapshot,
  WorkbookAutoFilterCustomCriterionSnapshot,
  WorkbookAutoFilterCustomOperator,
  WorkbookAutoFilterSnapshot,
  WorkbookAutoFilterValueCriteriaSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetAutoFilterTailElements = [
  'sortState',
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

function parseBooleanAttribute(value: unknown): boolean | undefined {
  if (value === true || value === '1' || value === 'true') {
    return true
  }
  if (value === false || value === '0' || value === 'false') {
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

function parseAutoFilterCustomOperator(value: unknown): WorkbookAutoFilterCustomOperator | undefined {
  switch (value) {
    case 'equal':
    case 'lessThan':
    case 'lessThanOrEqual':
    case 'notEqual':
    case 'greaterThanOrEqual':
    case 'greaterThan':
      return value
    default:
      return undefined
  }
}

function parseValueFilterCriteria(filters: Record<string, unknown>): WorkbookAutoFilterValueCriteriaSnapshot | null {
  const values = asArray(filters['filter']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['val'] !== 'string') {
      return []
    }
    return [entry['val']]
  })
  const blank = parseBooleanAttribute(filters['blank'])
  if (values.length === 0 && blank === undefined) {
    return null
  }
  return {
    ...(blank !== undefined ? { blank } : {}),
    values,
  }
}

function parseCustomFilterCriteria(customFilters: Record<string, unknown>): WorkbookAutoFilterCustomCriteriaSnapshot | null {
  const filters = asArray(customFilters['customFilter']).flatMap((entry): WorkbookAutoFilterCustomCriterionSnapshot[] => {
    if (!isRecord(entry) || typeof entry['val'] !== 'string') {
      return []
    }
    const operator = parseAutoFilterCustomOperator(entry['operator'])
    return [
      {
        ...(operator ? { operator } : {}),
        value: entry['val'],
      },
    ]
  })
  if (filters.length === 0) {
    return null
  }
  const and = parseBooleanAttribute(customFilters['and'])
  return {
    ...(and !== undefined ? { and } : {}),
    filters,
  }
}

function parseAutoFilterColumn(entry: unknown): WorkbookAutoFilterColumnSnapshot | null {
  if (!isRecord(entry)) {
    return null
  }
  const colId = Number(entry['colId'])
  if (!Number.isInteger(colId) || colId < 0) {
    return null
  }
  const filters = recordChild(entry, 'filters')
  const parsedFilters = filters ? parseValueFilterCriteria(filters) : null
  const customFilters = recordChild(entry, 'customFilters')
  const parsedCustomFilters = customFilters ? parseCustomFilterCriteria(customFilters) : null
  const hiddenButton = parseBooleanAttribute(entry['hiddenButton'])
  const showButton = parseBooleanAttribute(entry['showButton'])
  return {
    colId,
    ...(hiddenButton !== undefined ? { hiddenButton } : {}),
    ...(showButton !== undefined ? { showButton } : {}),
    ...(parsedFilters ? { filters: parsedFilters } : {}),
    ...(parsedCustomFilters ? { customFilters: parsedCustomFilters } : {}),
  }
}

function parseAutoFilter(sheetName: string, entry: unknown): WorkbookAutoFilterSnapshot | null {
  if (!isRecord(entry) || typeof entry['ref'] !== 'string') {
    return null
  }
  const range = parseRangeRef(sheetName, entry['ref'])
  if (!range) {
    return null
  }
  const criteria = asArray(entry['filterColumn']).flatMap((filterColumn) => {
    const parsed = parseAutoFilterColumn(filterColumn)
    return parsed ? [parsed] : []
  })
  return criteria.length > 0 ? { ...range, criteria } : range
}

export function readImportedSheetAutoFilters(sheetName: string, sheetXml: string): WorkbookAutoFilterSnapshot[] {
  if (!/<(?:[A-Za-z_][\w.-]*:)?autoFilter\b/u.test(sheetXml)) {
    return []
  }
  const parsed: unknown = xmlParser.parse(sheetXml)
  return readParsedWorksheetAutoFilters(sheetName, recordChild(parsed, 'worksheet'))
}

export function readImportedSheetAutoFiltersFromElementXml(sheetName: string, elementXml: string): WorkbookAutoFilterSnapshot[] {
  if (!/<(?:[A-Za-z_][\w.-]*:)?autoFilter\b/u.test(elementXml)) {
    return []
  }
  const parsed: unknown = xmlParser.parse(`<worksheet>${elementXml}</worksheet>`)
  return readParsedWorksheetAutoFilters(sheetName, recordChild(parsed, 'worksheet'))
}

function readParsedWorksheetAutoFilters(sheetName: string, worksheet: Record<string, unknown> | null): WorkbookAutoFilterSnapshot[] {
  return asArray(worksheet?.['autoFilter']).flatMap((entry) => {
    const filter = parseAutoFilter(sheetName, entry)
    return filter ? [filter] : []
  })
}

function valueFilterCriteriaXml(criteria: WorkbookAutoFilterValueCriteriaSnapshot): string {
  const attributes = criteria.blank !== undefined ? ` blank="${booleanAttribute(criteria.blank)}"` : ''
  const filters = criteria.values.map((value) => `<filter val="${escapeXml(value)}"/>`).join('')
  return filters ? `<filters${attributes}>${filters}</filters>` : `<filters${attributes}/>`
}

function customFilterCriteriaXml(criteria: WorkbookAutoFilterCustomCriteriaSnapshot): string {
  const attributes = criteria.and !== undefined ? ` and="${booleanAttribute(criteria.and)}"` : ''
  const filters = criteria.filters
    .map((filter) => {
      const operator = filter.operator ? ` operator="${escapeXml(filter.operator)}"` : ''
      return `<customFilter${operator} val="${escapeXml(filter.value)}"/>`
    })
    .join('')
  return filters ? `<customFilters${attributes}>${filters}</customFilters>` : `<customFilters${attributes}/>`
}

function autoFilterColumnXml(criteria: WorkbookAutoFilterColumnSnapshot): string {
  const attributes = [
    `colId="${String(criteria.colId)}"`,
    ...(criteria.hiddenButton !== undefined ? [`hiddenButton="${booleanAttribute(criteria.hiddenButton)}"`] : []),
    ...(criteria.showButton !== undefined ? [`showButton="${booleanAttribute(criteria.showButton)}"`] : []),
  ].join(' ')
  const children = [
    ...(criteria.filters ? [valueFilterCriteriaXml(criteria.filters)] : []),
    ...(criteria.customFilters ? [customFilterCriteriaXml(criteria.customFilters)] : []),
  ].join('')
  return children ? `<filterColumn ${attributes}>${children}</filterColumn>` : `<filterColumn ${attributes}/>`
}

function autoFilterXml(filter: WorkbookAutoFilterSnapshot, ref: string): string {
  const criteria = filter.criteria?.map(autoFilterColumnXml).join('') ?? ''
  return criteria ? `<autoFilter ref="${escapeXml(ref)}">${criteria}</autoFilter>` : `<autoFilter ref="${escapeXml(ref)}"/>`
}

function insertWorksheetAutoFilter(sheetXml: string, filter: WorkbookAutoFilterSnapshot, ref: string): string {
  const autoFilter = autoFilterXml(filter, ref)
  if (/<autoFilter\b/u.test(sheetXml)) {
    return sheetXml.replace(/<autoFilter\b[^>]*(?:\/>|>[\s\S]*?<\/autoFilter>)/u, autoFilter)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetAutoFilterTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${autoFilter}${sheetXml.slice(insertIndex)}`
}

export function addExportFiltersToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.filters ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const filter = (sheet.metadata?.filters ?? []).find((candidate) => candidate.sheetName === sheet.name)
      if (!filter) {
        return
      }
      const ref = rangeRefA1(filter)
      if (!ref) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetAutoFilter(sheetXml, filter, ref))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

export function readImportedWorkbookFilters(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookAutoFilterSnapshot[]> {
  const zip = readXlsxZipEntries(source)
  const filtersBySheet = new Map<string, WorkbookAutoFilterSnapshot[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<autoFilter\b/u.test(sheetXml)) {
      return
    }
    const autoFilters = readImportedSheetAutoFilters(sheetName, sheetXml)
    if (autoFilters.length > 0) {
      filtersBySheet.set(sheetName, autoFilters)
    }
  })

  return filtersBySheet
}
