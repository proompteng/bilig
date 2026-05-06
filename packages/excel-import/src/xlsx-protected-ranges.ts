import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'
import * as XLSX from 'xlsx'

import type { CellRangeRef, WorkbookRangeProtectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

const worksheetProtectedRangesTailElements = [
  'scenarios',
  'autoFilter',
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

function insertWorksheetProtectedRanges(sheetXml: string, rangeXml: readonly string[]): string {
  const protectedRanges = `<protectedRanges>${rangeXml.join('')}</protectedRanges>`
  if (/<protectedRanges\b/u.test(sheetXml)) {
    return sheetXml.replace(/<protectedRanges\b[^>]*(?:\/>|>[\s\S]*?<\/protectedRanges>)/u, protectedRanges)
  }

  let insertIndex = sheetXml.indexOf('</worksheet>')
  for (const elementName of worksheetProtectedRangesTailElements) {
    const elementIndex = sheetXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return sheetXml
  }
  return `${sheetXml.slice(0, insertIndex)}${protectedRanges}${sheetXml.slice(insertIndex)}`
}

function exportProtectedRangeXml(sheetName: string, protection: WorkbookRangeProtectionSnapshot): string | null {
  const name = protection.id.trim()
  const ref = protection.range.sheetName === sheetName ? rangeRefA1(protection.range) : null
  if (!name || !ref) {
    return null
  }
  return `<protectedRange name="${escapeXml(name)}" sqref="${escapeXml(ref)}"/>`
}

export function addExportProtectedRangesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => (sheet.metadata?.protectedRanges ?? []).length > 0)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const rangeXml = (sheet.metadata?.protectedRanges ?? []).flatMap((protection) => {
        const xml = exportProtectedRangeXml(sheet.name, protection)
        return xml ? [xml] : []
      })
      if (rangeXml.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetProtectedRanges(sheetXml, rangeXml))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

function uniqueProtectionId(input: { sheetName: string; name: unknown; range: CellRangeRef; usedIds: Set<string> }): string {
  const name = typeof input.name === 'string' ? input.name.trim() : ''
  const base = name || `protected-range:${input.sheetName}:${input.range.startAddress}:${input.range.endAddress}`
  let candidate = base
  let suffix = 2
  while (input.usedIds.has(candidate)) {
    candidate = `${base}:${String(suffix)}`
    suffix += 1
  }
  input.usedIds.add(candidate)
  return candidate
}

function parseProtectedRangeEntry(sheetName: string, entry: unknown, usedIds: Set<string>): WorkbookRangeProtectionSnapshot[] {
  if (!isRecord(entry) || typeof entry['sqref'] !== 'string') {
    return []
  }
  return entry['sqref'].split(/\s+/u).flatMap((rawRef) => {
    const ref = rawRef.trim()
    if (!ref) {
      return []
    }
    const range = parseRangeRef(sheetName, ref)
    if (!range) {
      return []
    }
    return [
      {
        id: uniqueProtectionId({ sheetName, name: entry['name'], range, usedIds }),
        range,
      },
    ]
  })
}

export function readImportedWorkbookProtectedRanges(
  bytes: Uint8Array,
  sheetNames: readonly string[],
): Map<string, WorkbookRangeProtectionSnapshot[]> {
  const zip = unzipSync(bytes)
  const protectedRangesBySheet = new Map<string, WorkbookRangeProtectionSnapshot[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const usedIds = new Set<string>()
    const protectedRanges = asArray(
      recordChild(recordChild(recordChild(parsed, 'worksheet'), 'protectedRanges'), 'protectedRange'),
    ).flatMap((entry) => parseProtectedRangeEntry(sheetName, entry, usedIds))
    if (protectedRanges.length > 0) {
      protectedRangesBySheet.set(sheetName, protectedRanges)
    }
  })

  return protectedRangesBySheet
}
