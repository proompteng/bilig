import { TextDecoder } from 'node:util'
import { inflateRawSync } from 'node:zlib'

import {
  getZipText,
  readXlsxZipEntriesLazyFromByteSource,
  readXlsxZipEntryMetadata,
  releaseLazyXlsxZipSource,
  type XlsxZipByteSource,
  type XlsxZipEntryMetadata,
} from '../packages/excel-import/src/xlsx-zip.js'
import type { WorkbookExternalWorkbookReferenceSnapshot } from '../packages/protocol/src/types.js'
import {
  addWorksheetFootprint,
  emptyWorksheetFootprint,
  inspectDeflatedWorksheetXmlBytes,
  inspectWorksheetXmlBytes,
  inspectWorksheetZipEntryFromLazyZip,
  type WorksheetFootprint,
  worksheetDimension,
} from './public-workbook-corpus-xlsx-worksheet-footprint.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

interface ZipEntryInfo {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
  readonly localHeaderOffset: number
}

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
  readonly targetMode?: string
}

interface ParsedSheet {
  readonly name: string
  readonly path: string | null
}

interface WorkbookPackageEntry {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
}

const decoder = new TextDecoder()
const eocdSignature = 0x06054b50
const centralDirectorySignature = 0x02014b50
const localFileHeaderSignature = 0x04034b50

export function isZipWorkbook(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04
}

export function inspectXlsxWorkbookFootprintLowMemory(bytes: Uint8Array, fileName: string): WorkbookFootprint {
  const context = readWorkbookPackageContext(bytes)
  const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
  const totals = emptyWorksheetFootprint()

  for (const sheet of context.sheets) {
    const worksheetBytes = sheet.path ? readZipEntryBytes(bytes, context.entriesByPath.get(sheet.path)) : null
    const footprint = worksheetBytes ? inspectWorksheetXmlBytes(worksheetBytes) : emptyWorksheetFootprint()
    addWorksheetFootprint(totals, footprint)
    dimensions.push(worksheetDimension(sheet.name, footprint))
  }

  return buildWorkbookFootprint(bytes, fileName, context, totals, dimensions)
}

export async function inspectXlsxWorkbookFootprintLowMemoryAsync(bytes: Uint8Array, fileName: string): Promise<WorkbookFootprint> {
  const context = readWorkbookPackageContext(bytes)
  const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
  const totals = emptyWorksheetFootprint()

  for (const sheet of context.sheets) {
    let footprint = emptyWorksheetFootprint()
    if (sheet.path) {
      // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential worksheet scanning keeps peak RSS bounded.
      footprint = await inspectWorksheetZipEntryLowMemory(bytes, context.entriesByPath.get(sheet.path))
    }
    addWorksheetFootprint(totals, footprint)
    dimensions.push(worksheetDimension(sheet.name, footprint))
  }

  return buildWorkbookFootprint(bytes, fileName, context, totals, dimensions)
}

export function inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source: XlsxZipByteSource, fileName: string): WorkbookFootprint | null {
  const zip = readXlsxZipEntriesLazyFromByteSource(source)
  const entryMetadata = readXlsxZipEntryMetadata(source)
  if (!zip || !entryMetadata) {
    return null
  }
  try {
    const context = readWorkbookPackageContextFromZip(zip, entryMetadata)
    const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
    const totals = emptyWorksheetFootprint()

    for (const sheet of context.sheets) {
      const footprint = sheet.path ? inspectWorksheetZipEntryFromLazyZip(zip, sheet.path) : emptyWorksheetFootprint()
      addWorksheetFootprint(totals, footprint)
      dimensions.push(worksheetDimension(sheet.name, footprint))
    }

    return buildWorkbookFootprintFromContext({
      fileName,
      context,
      totals,
      dimensions,
      readPartText: (path) => getZipText(zip, path),
    })
  } finally {
    releaseLazyXlsxZipSource(zip)
  }
}

function readWorkbookPackageContext(bytes: Uint8Array): {
  readonly entries: readonly WorkbookPackageEntry[]
  readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
  readonly workbookRelationships: readonly ParsedRelationship[]
  readonly workbookXml: string
  readonly sheets: readonly ParsedSheet[]
} {
  const entries = readZipCentralDirectory(bytes)
  const entriesByPath = new Map(entries.map((entry) => [entry.path, entry]))
  const workbookXml = readZipEntryText(bytes, entriesByPath.get('xl/workbook.xml')) ?? ''
  const workbookRelationships = parseRelationships(readZipEntryText(bytes, entriesByPath.get('xl/_rels/workbook.xml.rels')))
  const sheets = readWorkbookSheets(workbookXml, workbookRelationships, entries)
  return { entries, entriesByPath, workbookRelationships, workbookXml, sheets }
}

function readWorkbookPackageContextFromZip(
  zip: XlsxZipEntries,
  entryMetadata: readonly XlsxZipEntryMetadata[],
): {
  readonly entries: readonly WorkbookPackageEntry[]
  readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
  readonly workbookRelationships: readonly ParsedRelationship[]
  readonly workbookXml: string
  readonly sheets: readonly ParsedSheet[]
} {
  const entries = entryMetadata.map((entry) => ({
    path: entry.path,
    compressionMethod: entry.compressionMethod,
    compressedSize: entry.compressedSize,
  }))
  const entriesByPath = new Map(entries.map((entry) => [entry.path, entry]))
  const workbookXml = getZipText(zip, 'xl/workbook.xml') ?? ''
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const sheets = readWorkbookSheets(workbookXml, workbookRelationships, entries)
  return { entries, entriesByPath, workbookRelationships, workbookXml, sheets }
}

function buildWorkbookFootprint(
  bytes: Uint8Array,
  fileName: string,
  context: {
    readonly entries: readonly ZipEntryInfo[]
    readonly entriesByPath: ReadonlyMap<string, ZipEntryInfo>
    readonly workbookRelationships: readonly ParsedRelationship[]
    readonly workbookXml: string
    readonly sheets: readonly ParsedSheet[]
  },
  totals: WorksheetFootprint,
  dimensions: WorkbookFootprint['workbookMetadata']['dimensions'],
): WorkbookFootprint {
  return buildWorkbookFootprintFromContext({
    fileName,
    context,
    totals,
    dimensions,
    readPartText: (path) => readZipEntryText(bytes, context.entriesByPath.get(path)),
  })
}

function buildWorkbookFootprintFromContext(args: {
  readonly fileName: string
  readonly context: {
    readonly entries: readonly WorkbookPackageEntry[]
    readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
    readonly workbookRelationships: readonly ParsedRelationship[]
    readonly workbookXml: string
    readonly sheets: readonly ParsedSheet[]
  }
  readonly totals: WorksheetFootprint
  readonly dimensions: WorkbookFootprint['workbookMetadata']['dimensions']
  readonly readPartText: (path: string) => string | null
}): WorkbookFootprint {
  return {
    featureCounts: {
      sheetCount: args.context.sheets.length,
      cellCount: args.totals.cellCount,
      formulaCellCount: args.totals.formulaCellCount,
      valueCellCount: args.totals.valueCellCount,
      definedNameCount: countElementStarts(args.context.workbookXml, 'definedName'),
      tableCount: countZipEntries(args.context.entries, /^xl\/tables\/table[1-9][0-9]*\.xml$/u),
      chartCount: countZipEntries(args.context.entries, /^xl\/charts\/chart[1-9][0-9]*\.xml$/u),
      pivotCount: countZipEntries(args.context.entries, /^xl\/pivotTables\/pivotTable[1-9][0-9]*\.xml$/u),
      mergeCount: args.totals.mergeCount,
      styleRangeCount: 0,
      conditionalFormatCount: args.totals.conditionalFormatCount,
      dataValidationCount: args.totals.dataValidationCount,
      macroPayloadCount: args.context.entriesByPath.has('xl/vbaProject.bin') ? 1 : 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: args.fileName.replace(/\.(xlsx|xlsm|csv)$/iu, '') || args.fileName,
      sheetNames: args.context.sheets.map((sheet) => sheet.name),
      dimensions: args.dimensions,
    },
    externalWorkbookReferences: readExternalWorkbookReferences(
      args.readPartText,
      args.context.entries,
      args.context.workbookXml,
      args.context.workbookRelationships,
    ),
    largeSimpleXlsxImport: inspectLargeSimpleXlsxImportCompatibility(args.context, args.totals),
  }
}

function inspectLargeSimpleXlsxImportCompatibility(
  context: {
    readonly entries: readonly WorkbookPackageEntry[]
    readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
    readonly workbookXml: string
  },
  totals: WorksheetFootprint,
): WorkbookFootprint['largeSimpleXlsxImport'] {
  const blockers: string[] = []
  const unsupportedPackagePaths = context.entries.filter((entry) =>
    /^xl\/(?:charts|chartSheets|comments|ctrlProps|externalLinks|metadata\.xml|model|pivotCache|pivotTables|threadedComments|vbaProject\.bin)/u.test(
      entry.path,
    ),
  )
  if (unsupportedPackagePaths.length > 0) {
    blockers.push(`unsupported-package-parts=${String(unsupportedPackagePaths.length)}`)
  }
  const sharedStringsEntry = context.entriesByPath.get('xl/sharedStrings.xml')
  const hasSharedStrings = sharedStringsEntry !== undefined && sharedStringsEntry.compressedSize > 0
  if (definedNamesReferenceExternalWorkbook(context.workbookXml)) {
    blockers.push('external-defined-name-reference')
  }
  if (totals.largeSimpleUnsupportedDataValidationCount > 0) {
    blockers.push(`unsupported-data-validations=${String(totals.largeSimpleUnsupportedDataValidationCount)}`)
  }
  if (totals.largeSimpleUnsupportedFormulaCellCount > 0) {
    blockers.push(`unsupported-formula-cells=${String(totals.largeSimpleUnsupportedFormulaCellCount)}`)
  }
  if (totals.largeSimpleUnsupportedElementCount > 0) {
    blockers.push(`unsupported-worksheet-elements=${String(totals.largeSimpleUnsupportedElementCount)}`)
  }
  if (!hasSharedStrings) {
    if (totals.sharedStringCellCount > 0) {
      blockers.push(`missing-shared-strings=${String(totals.sharedStringCellCount)}`)
    }
    if (totals.valueCellCount === 0) {
      blockers.push('no-value-cells')
    }
    const blankCellCount = Math.max(0, totals.xmlCellCount - totals.valueCellCount)
    if (blankCellCount > 50_000 && blankCellCount > totals.valueCellCount * 8) {
      blockers.push(`style-only-blank-cells=${String(blankCellCount)}`)
    }
  }
  return {
    eligible: blockers.length === 0,
    blockers,
  }
}

function definedNamesReferenceExternalWorkbook(workbookXml: string): boolean {
  return /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?(?:^|[=,+(*/\s])'?\[[^\]]+\][\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?definedName>/u.test(
    workbookXml,
  )
}

function readZipCentralDirectory(bytes: Uint8Array): ZipEntryInfo[] {
  const eocdOffset = findEndOfCentralDirectory(bytes)
  const centralDirectorySize = readUint32(bytes, eocdOffset + 12)
  const centralDirectoryOffset = readUint32(bytes, eocdOffset + 16)
  if (centralDirectorySize === 0xffffffff || centralDirectoryOffset === 0xffffffff) {
    throw new Error('Zip64 central directories are not supported by the low-memory XLSX footprint scanner')
  }
  const entries: ZipEntryInfo[] = []
  let offset = centralDirectoryOffset
  const endOffset = centralDirectoryOffset + centralDirectorySize
  while (offset < endOffset) {
    if (readUint32(bytes, offset) !== centralDirectorySignature) {
      throw new Error('Invalid XLSX central directory entry')
    }
    const compressionMethod = readUint16(bytes, offset + 10)
    const compressedSize = readUint32(bytes, offset + 20)
    const fileNameLength = readUint16(bytes, offset + 28)
    const extraFieldLength = readUint16(bytes, offset + 30)
    const fileCommentLength = readUint16(bytes, offset + 32)
    const localHeaderOffset = readUint32(bytes, offset + 42)
    const path = normalizeZipPath(decodeBytes(bytes.subarray(offset + 46, offset + 46 + fileNameLength)))
    entries.push({ path, compressionMethod, compressedSize, localHeaderOffset })
    offset += 46 + fileNameLength + extraFieldLength + fileCommentLength
  }
  return entries
}

function findEndOfCentralDirectory(bytes: Uint8Array): number {
  const minOffset = Math.max(0, bytes.length - 65_557)
  for (let offset = bytes.length - 22; offset >= minOffset; offset -= 1) {
    if (readUint32(bytes, offset) === eocdSignature) {
      return offset
    }
  }
  throw new Error('Invalid XLSX zip: end of central directory not found')
}

function readZipEntryBytes(bytes: Uint8Array, entry: ZipEntryInfo | undefined): Uint8Array | null {
  const payload = readZipEntryPayload(bytes, entry)
  if (!payload) {
    return null
  }
  if (payload.compressionMethod === 0) {
    return payload.compressed
  }
  if (payload.compressionMethod === 8) {
    return inflateRawSync(payload.compressed)
  }
  throw new Error(`Unsupported XLSX zip compression method ${String(payload.compressionMethod)} for ${payload.path}`)
}

function readZipEntryPayload(
  bytes: Uint8Array,
  entry: ZipEntryInfo | undefined,
): { readonly compressed: Uint8Array; readonly compressionMethod: number; readonly path: string } | null {
  if (!entry) {
    return null
  }
  const offset = entry.localHeaderOffset
  if (readUint32(bytes, offset) !== localFileHeaderSignature) {
    throw new Error(`Invalid local file header for ${entry.path}`)
  }
  const fileNameLength = readUint16(bytes, offset + 26)
  const extraFieldLength = readUint16(bytes, offset + 28)
  const dataOffset = offset + 30 + fileNameLength + extraFieldLength
  const compressed = bytes.subarray(dataOffset, dataOffset + entry.compressedSize)
  return { compressed, compressionMethod: entry.compressionMethod, path: entry.path }
}

function readZipEntryText(bytes: Uint8Array, entry: ZipEntryInfo | undefined): string | null {
  const entryBytes = readZipEntryBytes(bytes, entry)
  return entryBytes ? decodeBytes(entryBytes) : null
}

function readWorkbookSheets(
  workbookXml: string,
  workbookRelationships: readonly ParsedRelationship[],
  entries: readonly WorkbookPackageEntry[],
): ParsedSheet[] {
  const worksheetRelationships = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type.endsWith('/worksheet') || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, resolveTargetPath('xl/workbook.xml', relationship.target)]),
  )
  const sheets = elementTags(workbookXml, 'sheet').map((tag, index) => {
    const name = readXmlAttribute(tag, 'name') ?? `Sheet${String(index + 1)}`
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    return {
      name,
      path: relationshipId ? (worksheetRelationships.get(relationshipId) ?? null) : null,
    }
  })
  if (sheets.length > 0) {
    return sheets
  }
  return entries
    .map((entry) => /^xl\/worksheets\/sheet([1-9][0-9]*)\.xml$/u.exec(entry.path)?.[1])
    .flatMap((sheetNumber) => (sheetNumber ? [{ name: `Sheet${sheetNumber}`, path: `xl/worksheets/sheet${sheetNumber}.xml` }] : []))
    .toSorted((left, right) => left.name.localeCompare(right.name, 'en-US', { numeric: true }))
}

async function inspectWorksheetZipEntryLowMemory(bytes: Uint8Array, entry: ZipEntryInfo | undefined): Promise<WorksheetFootprint> {
  const payload = readZipEntryPayload(bytes, entry)
  if (!payload) {
    return emptyWorksheetFootprint()
  }
  if (payload.compressionMethod === 0) {
    return inspectWorksheetXmlBytes(payload.compressed)
  }
  if (payload.compressionMethod === 8) {
    return inspectDeflatedWorksheetXmlBytes(payload.compressed)
  }
  throw new Error(`Unsupported XLSX zip compression method ${String(payload.compressionMethod)} for ${payload.path}`)
}

function readExternalWorkbookReferences(
  readPartText: (path: string) => string | null,
  entries: readonly WorkbookPackageEntry[],
  workbookXml: string,
  workbookRelationships: readonly ParsedRelationship[],
): readonly WorkbookExternalWorkbookReferenceSnapshot[] {
  const workbookTargets = readWorkbookExternalLinkTargets(workbookXml, workbookRelationships)
  const linkTargets = workbookTargets.size > 0 ? workbookTargets : readFallbackExternalLinkTargets(entries)
  return [...linkTargets.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .flatMap(([bookIndex, path]) => {
      const xml = readPartText(path)
      if (!xml) {
        return []
      }
      const relationships = parseRelationships(readPartText(externalLinkRelationshipsPartPath(path)))
      const externalBookRelationshipId =
        readXmlAttribute(elementTags(xml, 'externalBook')[0] ?? '', 'r:id') ??
        readXmlAttribute(elementTags(xml, 'externalBook')[0] ?? '', 'id')
      const linkedWorkbookRelationship =
        (externalBookRelationshipId ? relationships.find((relationship) => relationship.id === externalBookRelationshipId) : undefined) ??
        relationships.find((relationship) => relationship.type.endsWith('/externalLinkPath'))
      const target = linkedWorkbookRelationship?.target
      const workbookName = target ? workbookNameFromExternalTarget(target) : undefined
      const sheetNames = elementTags(xml, 'sheetName').flatMap((tag) => {
        const name = readXmlAttribute(tag, 'val')
        return name ? [name] : []
      })
      const reference: WorkbookExternalWorkbookReferenceSnapshot = {
        bookIndex,
        packagePath: path,
        ...(target ? { target } : {}),
        ...(linkedWorkbookRelationship?.targetMode ? { targetMode: linkedWorkbookRelationship.targetMode } : {}),
        ...(workbookName ? { workbookName } : {}),
        ...(sheetNames.length > 0 ? { sheetNames } : {}),
      }
      return [reference]
    })
}

function readWorkbookExternalLinkTargets(workbookXml: string, workbookRelationships: readonly ParsedRelationship[]): Map<number, string> {
  const targets = new Map<number, string>()
  let bookIndex = 1
  for (const tag of elementTags(workbookXml, 'externalReference')) {
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    const relationship = relationshipId
      ? workbookRelationships.find((candidate) => candidate.id === relationshipId && candidate.type.endsWith('/externalLink'))
      : undefined
    if (relationship) {
      targets.set(bookIndex, resolveTargetPath('xl/workbook.xml', relationship.target))
    }
    bookIndex += 1
  }
  return targets
}

function readFallbackExternalLinkTargets(entries: readonly ZipEntryInfo[]): Map<number, string> {
  const targets = new Map<number, string>()
  for (const entry of entries) {
    const match = /^xl\/externalLinks\/externalLink([1-9][0-9]*)\.xml$/u.exec(entry.path)
    if (match) {
      targets.set(Number(match[1]), entry.path)
    }
  }
  return targets
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  return elementTags(xml, 'Relationship').flatMap((tag) => {
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    if (!id || !target || !type) {
      return []
    }
    const targetMode = readXmlAttribute(tag, 'TargetMode')
    return [{ id, target, type, ...(targetMode ? { targetMode } : {}) }]
  })
}

function externalLinkRelationshipsPartPath(partPath: string): string {
  const fileName = partPath.slice(partPath.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function resolveTargetPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return normalizeZipPath(parts.join('/'))
}

function workbookNameFromExternalTarget(target: string): string | undefined {
  const targetPath = target.split(/[?#]/u)[0] ?? target
  const normalizedTarget = targetPath.replace(/^file:\/+/iu, '').replace(/\\/gu, '/')
  const lastSegment = normalizedTarget
    .split('/')
    .toReversed()
    .find((segment) => segment.length > 0)
  if (!lastSegment) {
    return undefined
  }
  try {
    return decodeURIComponent(lastSegment)
  } catch {
    return lastSegment
  }
}

function countZipEntries(entries: readonly WorkbookPackageEntry[], pattern: RegExp): number {
  return entries.filter((entry) => pattern.test(entry.path)).length
}

function elementTags(xml: string, elementName: string): string[] {
  const tags: string[] = []
  let index = 0
  while (index < xml.length) {
    const tagStart = xml.indexOf(`<${elementName}`, index)
    if (tagStart < 0) {
      break
    }
    if (!isElementStart(xml, tagStart, elementName)) {
      index = tagStart + elementName.length + 1
      continue
    }
    const tagEnd = xml.indexOf('>', tagStart + elementName.length + 1)
    if (tagEnd < 0) {
      break
    }
    tags.push(xml.slice(tagStart, tagEnd + 1))
    index = tagEnd + 1
  }
  return tags
}

function countElementStarts(xml: string, elementName: string): number {
  let count = 0
  let index = 0
  while (index < xml.length) {
    const tagStart = xml.indexOf(`<${elementName}`, index)
    if (tagStart < 0) {
      return count
    }
    if (isElementStart(xml, tagStart, elementName)) {
      count += 1
    }
    index = tagStart + elementName.length + 1
  }
  return count
}

function isElementStart(xml: string, index: number, elementName: string): boolean {
  if (!xml.startsWith(`<${elementName}`, index)) {
    return false
  }
  const next = xml[index + elementName.length + 1]
  return next === undefined || /\s|\/|>/u.test(next)
}

function readXmlAttribute(tag: string, name: string): string | null {
  return readXmlAttributeInRange(tag, 0, tag.length, name)
}

function readXmlAttributeInRange(source: string, start: number, end: number, name: string): string | null {
  const pattern = new RegExp(`\\s${escapeRegExp(name)}=(["'])([\\s\\S]*?)\\1`, 'u')
  const match = pattern.exec(source.slice(start, end))
  return match?.[2] ? decodeXmlEntities(match[2]) : null
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(#x?[0-9a-f]+|amp|apos|gt|lt|quot);/giu, (entity, raw: string) => {
    switch (raw) {
      case 'amp':
        return '&'
      case 'apos':
        return "'"
      case 'gt':
        return '>'
      case 'lt':
        return '<'
      case 'quot':
        return '"'
      default: {
        const radix = raw.toLowerCase().startsWith('#x') ? 16 : 10
        const digits = raw.replace(/^#x?/iu, '')
        const codePoint = Number.parseInt(digits, radix)
        return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity
      }
    }
  })
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function decodeBytes(bytes: Uint8Array): string {
  return decoder.decode(bytes)
}

function readUint16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8)
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
