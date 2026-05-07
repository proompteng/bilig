import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type { WorkbookSheetTabColorSnapshot, WorkbookSnapshot } from '@bilig/protocol'

type ZipEntries = Record<string, Uint8Array>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  removeNSPrefix: true,
})

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
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

function readAttribute(value: unknown): string | undefined {
  const text = typeof value === 'number' || typeof value === 'boolean' ? String(value) : typeof value === 'string' ? value.trim() : ''
  return text.length > 0 ? text : undefined
}

function readRgbAttribute(value: unknown): string | undefined {
  const text = readAttribute(value)?.toUpperCase()
  return text && /^[0-9A-F]{6}([0-9A-F]{2})?$/u.test(text) ? text : undefined
}

function normalizeTabColor(tabColor: WorkbookSheetTabColorSnapshot | undefined): WorkbookSheetTabColorSnapshot | null {
  if (!tabColor) {
    return null
  }
  const normalized: WorkbookSheetTabColorSnapshot = {}
  const rgb = readRgbAttribute(tabColor.rgb)
  if (rgb !== undefined) {
    normalized.rgb = rgb
  }
  const theme = readAttribute(tabColor.theme)
  if (theme !== undefined) {
    normalized.theme = theme
  }
  const tint = readAttribute(tabColor.tint)
  if (tint !== undefined) {
    normalized.tint = tint
  }
  const indexed = readAttribute(tabColor.indexed)
  if (indexed !== undefined) {
    normalized.indexed = indexed
  }
  const auto = readAttribute(tabColor.auto)
  if (auto !== undefined) {
    normalized.auto = auto
  }
  return Object.keys(normalized).length > 0 ? normalized : null
}

function tabColorFromRecord(record: Record<string, unknown>): WorkbookSheetTabColorSnapshot | null {
  const rgb = readRgbAttribute(record['rgb'])
  const theme = readAttribute(record['theme'])
  const tint = readAttribute(record['tint'])
  const indexed = readAttribute(record['indexed'])
  const auto = readAttribute(record['auto'])
  return normalizeTabColor({
    ...(rgb !== undefined ? { rgb } : {}),
    ...(theme !== undefined ? { theme } : {}),
    ...(tint !== undefined ? { tint } : {}),
    ...(indexed !== undefined ? { indexed } : {}),
    ...(auto !== undefined ? { auto } : {}),
  })
}

function tabColorXml(tabColor: WorkbookSheetTabColorSnapshot): string {
  const attributes: string[] = []
  if (tabColor.rgb !== undefined) {
    attributes.push(`rgb="${escapeXml(tabColor.rgb)}"`)
  }
  if (tabColor.theme !== undefined) {
    attributes.push(`theme="${escapeXml(tabColor.theme)}"`)
  }
  if (tabColor.tint !== undefined) {
    attributes.push(`tint="${escapeXml(tabColor.tint)}"`)
  }
  if (tabColor.indexed !== undefined) {
    attributes.push(`indexed="${escapeXml(tabColor.indexed)}"`)
  }
  if (tabColor.auto !== undefined) {
    attributes.push(`auto="${escapeXml(tabColor.auto)}"`)
  }
  return `<tabColor ${attributes.join(' ')}/>`
}

function removeExistingTabColor(sheetPrInnerXml: string): string {
  return sheetPrInnerXml.replace(/<tabColor\b[^>]*(?:\/>|>[\s\S]*?<\/tabColor>)/u, '')
}

function addTabColorToSheetPr(sheetPrXml: string, tabColor: WorkbookSheetTabColorSnapshot): string {
  const colorXml = tabColorXml(tabColor)
  const selfClosingMatch = /^<sheetPr\b([^>]*)\/>$/u.exec(sheetPrXml)
  if (selfClosingMatch) {
    return `<sheetPr${selfClosingMatch[1] ?? ''}>${colorXml}</sheetPr>`
  }

  const expandedMatch = /^<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>$/u.exec(sheetPrXml)
  if (!expandedMatch) {
    return sheetPrXml
  }
  return `<sheetPr${expandedMatch[1] ?? ''}>${colorXml}${removeExistingTabColor(expandedMatch[2] ?? '')}</sheetPr>`
}

function insertWorksheetTabColor(sheetXml: string, tabColor: WorkbookSheetTabColorSnapshot): string {
  const sheetPrMatch = /<sheetPr\b[^>]*(?:\/>|>[\s\S]*?<\/sheetPr>)/u.exec(sheetXml)
  if (sheetPrMatch) {
    return sheetXml.replace(sheetPrMatch[0], addTabColorToSheetPr(sheetPrMatch[0], tabColor))
  }
  return sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1><sheetPr>${tabColorXml(tabColor)}</sheetPr>`)
}

export function addExportSheetTabColorsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  if (!snapshot.sheets.some((sheet) => normalizeTabColor(sheet.metadata?.tabColor) !== null)) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const tabColor = normalizeTabColor(sheet.metadata?.tabColor)
      if (!tabColor) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      setZipText(zip, sheetPath, insertWorksheetTabColor(sheetXml, tabColor))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}

export function readImportedWorkbookSheetTabColors(
  bytes: Uint8Array,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetTabColorSnapshot> {
  const zip = unzipSync(bytes)
  const tabColorsBySheet = new Map<string, WorkbookSheetTabColorSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetXml = getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)
    if (!sheetXml || !/<tabColor\b/u.test(sheetXml)) {
      return
    }
    const parsed: unknown = xmlParser.parse(sheetXml)
    const tabColor = recordChild(recordChild(recordChild(parsed, 'worksheet'), 'sheetPr'), 'tabColor')
    if (!tabColor) {
      return
    }
    const snapshot = tabColorFromRecord(tabColor)
    if (snapshot) {
      tabColorsBySheet.set(sheetName, snapshot)
    }
  })

  return tabColorsBySheet
}
