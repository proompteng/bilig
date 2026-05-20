import { normalizeWorkbookName } from './workbook-import-helpers.js'
import type {
  LargeSimpleXlsxImportStats,
  LargeSimpleXlsxOwnedSourceReleaseEvidence,
  LargeSimpleXlsxSheetDimension,
} from './xlsx-large-simple-import.js'
import { LargeSimpleXlsxImportPhaseRecorder } from './xlsx-large-simple-import-telemetry.js'
import { parseHeadlessLargeSimpleWorksheetFromChunks } from './xlsx-large-simple-headless-worksheet-scanner.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  normalizeZipPath,
  readLazyXlsxZipSourceByteLength,
  releaseLazyXlsxZipSource,
  type XlsxZipEntries,
} from './xlsx-zip.js'

export interface LargeSimpleXlsxHeadlessInspectResult {
  readonly workbookName: string
  readonly sheetNames: string[]
  readonly warnings: string[]
  readonly stats: LargeSimpleXlsxImportStats
}

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

interface WorkbookRelationship {
  readonly id: string
  readonly type: string
  readonly target: string
}

const defaultLargeSimpleXlsxByteThreshold = 1_000_000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const sharedStringsPath = 'xl/sharedStrings.xml'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const headlessZipEntryChunkSize = 16 * 1024
const definedNameElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?definedName)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const unsupportedPackagePathPattern =
  /^xl\/(?:charts|chartSheets|comments|ctrlProps|externalLinks|model|pivotCache|pivotTables|threadedComments|vbaProject\.bin)/u

export function tryInspectLargeSimpleXlsxHeadless(
  source: { readonly byteLength: number },
  fileName: string,
  zip: XlsxZipEntries,
  options: {
    readonly afterWorksheetScan?: () => void
    readonly minByteLength?: number
    readonly releaseZipSource?: boolean
    readonly releaseOwnedSourceBytes?: () => LargeSimpleXlsxOwnedSourceReleaseEvidence | undefined
  } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  if (source.byteLength < (options.minByteLength ?? defaultLargeSimpleXlsxByteThreshold)) {
    return null
  }
  const phaseRecorder = new LargeSimpleXlsxImportPhaseRecorder()
  const zipSetupStart = phaseRecorder.start()
  const packagePaths = Object.keys(zip).map(normalizeZipPath)
  if (packagePaths.some((path) => unsupportedPackagePathPattern.test(path))) {
    return null
  }
  const workbookXml = getZipText(zip, workbookPath)
  const workbookRelationshipsXml = getZipText(zip, workbookRelationshipsPath)
  if (!workbookXml || !workbookRelationshipsXml) {
    return null
  }
  const workbookSheets = readWorkbookSheets(workbookXml)
  const worksheetPathsByRelationshipId = readWorksheetPathsByRelationshipId(workbookRelationshipsXml)
  const definedNames = inspectWorkbookDefinedNames(
    workbookXml,
    workbookSheets.map((entry) => entry.name),
  )
  if (workbookSheets.length === 0 || worksheetPathsByRelationshipId.size === 0 || definedNames.externalWorkbookReferenceSeen) {
    return null
  }
  const worksheetEntries = workbookSheets.flatMap((entry) => {
    const path = worksheetPathsByRelationshipId.get(entry.relationshipId)
    return path ? [{ name: entry.name, path }] : []
  })
  if (worksheetEntries.length !== workbookSheets.length) {
    return null
  }
  const hasSharedStrings = packagePaths.includes(sharedStringsPath)
  delete zip[workbookPath]
  delete zip[workbookRelationshipsPath]
  phaseRecorder.finish('zip-setup', zipSetupStart)

  const dimensions: LargeSimpleXlsxSheetDimension[] = []
  let cellCount = 0
  let formulaCellCount = 0
  let valueCellCount = 0
  let tableCount = 0
  let mergeCount = 0
  let conditionalFormatCount = 0
  let dataValidationCount = 0
  for (const [order, entry] of worksheetEntries.entries()) {
    const worksheetScanStart = phaseRecorder.start()
    const scan = parseHeadlessLargeSimpleWorksheetFromChunks(
      (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk, { chunkSize: headlessZipEntryChunkSize }),
      order,
      { hasSharedStrings, sheetName: entry.name },
    )
    if (!scan || (!hasSharedStrings && scan.valueCellCount === 0)) {
      return null
    }
    delete zip[entry.path]
    cellCount += scan.cellCount
    formulaCellCount += scan.formulaCellCount
    valueCellCount += scan.valueCellCount
    tableCount += scan.tableCount ?? 0
    mergeCount += scan.mergeCount ?? 0
    conditionalFormatCount += scan.conditionalFormatCount ?? 0
    dataValidationCount += scan.dataValidationCount
    dimensions.push({
      sheetName: entry.name,
      rowCount: scan.rowCount,
      columnCount: scan.columnCount,
      nonEmptyCellCount: scan.cellCount,
      usedRange: scan.usedRange,
    })
    options.afterWorksheetScan?.()
    phaseRecorder.finish('worksheet-scan', worksheetScanStart)
  }
  delete zip[sharedStringsPath]
  if (options.releaseZipSource === true) {
    const zipSourceReleaseStart = phaseRecorder.start()
    const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
    releaseLazyXlsxZipSource(zip)
    const ownedSourceReleaseEvidence = options.releaseOwnedSourceBytes?.()
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
      ...ownedSourceReleaseEvidence,
    })
  }
  return {
    workbookName: normalizeWorkbookName(fileName),
    sheetNames: workbookSheets.map((entry) => entry.name),
    warnings: definedNames.ignoredCount > 0 ? ['Some defined names were ignored during XLSX import.'] : [],
    stats: {
      sheetCount: workbookSheets.length,
      cellCount,
      formulaCellCount,
      valueCellCount,
      definedNameCount: definedNames.count,
      tableCount,
      mergeCount,
      conditionalFormatCount,
      dataValidationCount,
      warningCount: definedNames.ignoredCount > 0 ? 1 : 0,
      dimensions,
      phaseTelemetry: phaseRecorder.entries(),
    },
  }
}

function readWorkbookSheets(workbookXml: string): WorkbookSheetEntry[] {
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const name = readXmlAttribute(tag, 'name')
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    return name && relationshipId ? [{ name: decodeXmlText(name), relationshipId }] : []
  })
}

function readWorksheetPathsByRelationshipId(workbookRelationshipsXml: string): Map<string, string> {
  return new Map(
    readRelationships(workbookRelationshipsXml).flatMap((relationship) => {
      if (relationship.type !== worksheetRelationshipType && !relationship.target.includes('worksheets/')) {
        return []
      }
      return [[relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]]
    }),
  )
}

function readRelationships(relationshipsXml: string): WorkbookRelationship[] {
  return [...relationshipsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const type = readXmlAttribute(tag, 'Type')
    const target = readXmlAttribute(tag, 'Target')
    return id && type && target ? [{ id, type, target }] : []
  })
}

function inspectWorkbookDefinedNames(
  workbookXml: string,
  sheetNames: readonly string[],
): {
  readonly count: number
  readonly externalWorkbookReferenceSeen: boolean
  readonly ignoredCount: number
} {
  let count = 0
  let ignoredCount = 0
  let externalWorkbookReferenceSeen = false
  for (const match of workbookXml.matchAll(definedNameElementPattern)) {
    const xml = match[0]
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xml)?.[0]
    const name = openingTag ? readXmlAttribute(openingTag, 'name')?.trim() : ''
    const localSheetId = openingTag ? readNonNegativeIntegerAttribute(openingTag, 'localSheetId') : null
    const rawValue = openingTag?.endsWith('/>') ? '' : decodeXmlText(xml.replace(/^<[^>]*>/u, '').replace(/<\/[^>]*>$/u, '')).trim()
    if (!name || rawValue.length === 0 || (localSheetId !== null && sheetNames[localSheetId] === undefined)) {
      ignoredCount += 1
      continue
    }
    if (definedNameReferencesExternalWorkbook(rawValue)) {
      externalWorkbookReferenceSeen = true
      continue
    }
    count += 1
  }
  return { count, externalWorkbookReferenceSeen, ignoredCount }
}

function definedNameReferencesExternalWorkbook(value: string): boolean {
  return /(?:^|[=,+(*/\s])'?\[[^\]]+\]/u.test(value)
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readNonNegativeIntegerAttribute(xml: string, attributeName: string): number | null {
  const raw = readXmlAttribute(xml, attributeName)
  const value = raw === null || raw.trim().length === 0 ? Number.NaN : Number(raw)
  return Number.isInteger(value) && value >= 0 ? value : null
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function resolveTargetPath(basePath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const parts = basePath.split('/')
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
