import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type {
  WorkbookPackageRelationshipSnapshot,
  WorkbookPivotArtifactsSnapshot,
  WorkbookPivotPackagePartSnapshot,
  WorkbookSheetPivotArtifactsSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, type XlsxZipEntries } from './xlsx-zip.js'

export interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
  readonly targetMode?: string
}

export const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const pivotTableRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
export const pivotCacheDefinitionRelationshipType =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
export const pivotCacheRecordsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'
export const pivotTableContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'
export const pivotCacheDefinitionContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
export const pivotCacheRecordsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const pivotPackagePartPathPattern = /^xl\/(?:pivotTables|pivotCache)\/.+/u
const pivotTableDefinitionElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?pivotTableDefinition\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pivotTableDefinition>)/gu

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

export function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

export function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

export function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  const parsed: unknown = xmlParser.parse(xml)
  return asArray(recordChild(parsed, 'Relationships')?.['Relationship']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['Id'] !== 'string' || typeof entry['Target'] !== 'string' || typeof entry['Type'] !== 'string') {
      return []
    }
    return [
      {
        id: entry['Id'],
        target: entry['Target'],
        type: entry['Type'],
        ...(typeof entry['TargetMode'] === 'string' ? { targetMode: entry['TargetMode'] } : {}),
      },
    ]
  })
}

export function nextRelationshipId(relationships: readonly ParsedRelationship[]): string {
  let next = 1
  for (const relationship of relationships) {
    const match = /^rId(\d+)$/u.exec(relationship.id)
    if (match) {
      next = Math.max(next, Number(match[1]) + 1)
    }
  }
  return `rId${String(next)}`
}

export function buildRelationshipsXml(relationships: readonly ParsedRelationship[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) =>
        `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(
          relationship.target,
        )}"${relationship.targetMode ? ` TargetMode="${escapeXml(relationship.targetMode)}"` : ''}/>`,
    ),
    '</Relationships>',
  ].join('')
}

export function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (contentTypesXml.includes(`PartName="${partName}"`)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`)
}

export function resolveTargetPath(basePartPath: string, target: string): string {
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
  return parts.join('/')
}

export function ensureRelationshipNamespace(xml: string): string {
  if (/xmlns:r=/u.test(xml)) {
    return xml
  }
  return xml.replace(/<([A-Za-z0-9:]+)\b([^>]*)>/u, `<$1$2 xmlns:r="${officeRelationshipNamespace}">`)
}

function relationshipSnapshot(relationship: ParsedRelationship): WorkbookPackageRelationshipSnapshot {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function parsedRelationship(relationship: WorkbookPackageRelationshipSnapshot): ParsedRelationship {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function addRelationshipsWithFreshIds(
  relationships: ParsedRelationship[],
  additions: readonly WorkbookPackageRelationshipSnapshot[] | undefined,
): Map<string, string> {
  const idMap = new Map<string, string>()
  for (const addition of additions ?? []) {
    const nextId = nextRelationshipId(relationships)
    relationships.push({ ...parsedRelationship(addition), id: nextId })
    idMap.set(addition.id, nextId)
  }
  return idMap
}

function replaceRelationshipIds(xml: string, relationshipIds: ReadonlyMap<string, string>): string {
  if (relationshipIds.size === 0) {
    return xml
  }
  return xml.replace(/\b(?:r:)?id=(["'])([\s\S]*?)\1/gu, (match, quote: string, id: string) => {
    const nextId = relationshipIds.get(id)
    return nextId ? match.replace(`${quote}${id}${quote}`, `${quote}${nextId}${quote}`) : match
  })
}

function pivotPartContentType(path: string): string | null {
  if (/^xl\/pivotTables\/pivotTable\d+\.xml$/u.test(path)) {
    return pivotTableContentType
  }
  if (/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/u.test(path)) {
    return pivotCacheDefinitionContentType
  }
  if (/^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/u.test(path)) {
    return pivotCacheRecordsContentType
  }
  return null
}

function addPivotPartContentTypes(contentTypesXml: string, parts: readonly WorkbookPivotPackagePartSnapshot[]): string {
  let output = contentTypesXml
  for (const part of parts) {
    const contentType = pivotPartContentType(part.path)
    if (contentType) {
      output = addContentTypeOverride(output, `/${part.path}`, contentType)
    }
  }
  return output
}

function workbookPivotCachesXml(workbookXml: string | null): string | undefined {
  return workbookXml?.match(/<(?:[A-Za-z_][\w.-]*:)?pivotCaches\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pivotCaches>)/u)?.[0]
}

function insertWorkbookPivotCachesXml(workbookXml: string, pivotCachesXml: string, relationshipIds: ReadonlyMap<string, string>): string {
  const nextPivotCachesXml = replaceRelationshipIds(pivotCachesXml, relationshipIds)
  const withoutPivotCaches = ensureRelationshipNamespace(
    workbookXml.replace(/<(?:[A-Za-z_][\w.-]*:)?pivotCaches\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?pivotCaches>)/u, ''),
  )
  if (withoutPivotCaches.includes('</calcPr>')) {
    return withoutPivotCaches.replace('</calcPr>', `</calcPr>${nextPivotCachesXml}`)
  }
  if (withoutPivotCaches.includes('</definedNames>')) {
    return withoutPivotCaches.replace('</definedNames>', `</definedNames>${nextPivotCachesXml}`)
  }
  if (withoutPivotCaches.includes('</sheets>')) {
    return withoutPivotCaches.replace('</sheets>', `</sheets>${nextPivotCachesXml}`)
  }
  return withoutPivotCaches.replace('</workbook>', `${nextPivotCachesXml}</workbook>`)
}

class LazyPivotPackagePartSnapshot implements WorkbookPivotPackagePartSnapshot {
  declare readonly xml: string
  private xmlCache: string | undefined

  constructor(
    readonly path: string,
    private readonly readXml: () => string,
  ) {
    Object.defineProperty(this, 'xml', {
      configurable: true,
      enumerable: true,
      get: () => this.getXml(),
    })
  }

  private getXml(): string {
    this.xmlCache ??= this.readXml()
    return this.xmlCache
  }
}

function lazyPivotPackagePartSnapshot(path: string, readXml: () => string): WorkbookPivotPackagePartSnapshot {
  return new LazyPivotPackagePartSnapshot(path, readXml)
}

function readPivotPackageParts(zip: XlsxZipEntries): WorkbookPivotPackagePartSnapshot[] {
  return Object.keys(zip)
    .filter((path) => pivotPackagePartPathPattern.test(path))
    .toSorted()
    .map((path) =>
      lazyPivotPackagePartSnapshot(path, () => {
        const bytes = zip[path]
        if (bytes) {
          Reflect.deleteProperty(zip, path)
        }
        return bytes ? strFromU8(bytes) : ''
      }),
    )
}

function readPivotTableDefinitionsXml(sheetXml: string | null): string | undefined {
  if (!sheetXml) {
    return undefined
  }
  pivotTableDefinitionElementPattern.lastIndex = 0
  const definitions = [...sheetXml.matchAll(pivotTableDefinitionElementPattern)].map((match) => match[0])
  return definitions.length > 0 ? definitions.join('') : undefined
}

function insertPivotTableDefinitionsXml(
  sheetXml: string,
  pivotTableDefinitionsXml: string | undefined,
  relationshipIds: ReadonlyMap<string, string>,
): string {
  if (!pivotTableDefinitionsXml) {
    return sheetXml
  }
  pivotTableDefinitionElementPattern.lastIndex = 0
  const withoutPivotTableDefinitions = ensureRelationshipNamespace(sheetXml.replace(pivotTableDefinitionElementPattern, ''))
  const nextDefinitionsXml = replaceRelationshipIds(pivotTableDefinitionsXml, relationshipIds)
  return withoutPivotTableDefinitions.replace('</worksheet>', `${nextDefinitionsXml}</worksheet>`)
}

export function readImportedPivotArtifacts(
  zip: XlsxZipEntries,
  sheetNames: readonly string[],
  options: {
    readonly readWorksheetPivotTableDefinitionsXml?: boolean
    readonly worksheetPivotTableDefinitionsXmlByName?: ReadonlyMap<string, string>
  } = {},
): {
  readonly artifacts: WorkbookPivotArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetPivotArtifactsSnapshot>
} {
  const parts = readPivotPackageParts(zip)
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
    .filter((relationship) => relationship.type === pivotCacheDefinitionRelationshipType)
    .map(relationshipSnapshot)
  const pivotCachesXml = workbookPivotCachesXml(getZipText(zip, 'xl/workbook.xml'))
  const sheetArtifactsByName = new Map<string, WorkbookSheetPivotArtifactsSnapshot>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const relationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
      .filter((relationship) => relationship.type === pivotTableRelationshipType)
      .map(relationshipSnapshot)
    const pivotTableDefinitionsXml =
      options.worksheetPivotTableDefinitionsXmlByName?.get(sheetName) ??
      (options.readWorksheetPivotTableDefinitionsXml === false
        ? undefined
        : readPivotTableDefinitionsXml(getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`)))
    if (relationships.length > 0 || pivotTableDefinitionsXml) {
      sheetArtifactsByName.set(sheetName, {
        relationships,
        ...(pivotTableDefinitionsXml ? { pivotTableDefinitionsXml } : {}),
      })
    }
  })

  const artifacts =
    parts.length > 0 || workbookRelationships.length > 0 || pivotCachesXml
      ? {
          parts,
          ...(pivotCachesXml ? { workbookPivotCachesXml: pivotCachesXml } : {}),
          ...(workbookRelationships.length > 0 ? { workbookRelationships } : {}),
        }
      : undefined
  return { artifacts, sheetArtifactsByName }
}

export function addExportPreservedPivotArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const artifacts = snapshot.workbook.metadata?.pivotArtifacts
  if (!artifacts) {
    return bytes
  }
  const zip = unzipSync(bytes)
  let changed = false

  for (const part of artifacts.parts) {
    if (!pivotPackagePartPathPattern.test(part.path)) {
      continue
    }
    setZipText(zip, part.path, part.xml)
    changed = true
  }

  const contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  if (contentTypesXml.length > 0) {
    const nextContentTypesXml = addPivotPartContentTypes(contentTypesXml, artifacts.parts)
    if (nextContentTypesXml !== contentTypesXml) {
      setZipText(zip, '[Content_Types].xml', nextContentTypesXml)
      changed = true
    }
  }

  const workbookRelationshipPath = 'xl/_rels/workbook.xml.rels'
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipPath))
  const workbookRelationshipIds = addRelationshipsWithFreshIds(workbookRelationships, artifacts.workbookRelationships)
  if ((artifacts.workbookRelationships?.length ?? 0) > 0) {
    setZipText(zip, workbookRelationshipPath, buildRelationshipsXml(workbookRelationships))
    changed = true
  }

  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (workbookXml && artifacts.workbookPivotCachesXml) {
    setZipText(zip, 'xl/workbook.xml', insertWorkbookPivotCachesXml(workbookXml, artifacts.workbookPivotCachesXml, workbookRelationshipIds))
    changed = true
  }

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetArtifacts = sheet.metadata?.pivotArtifacts
      if (!sheetArtifacts) {
        return
      }
      const relationshipPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const relationships = parseRelationships(getZipText(zip, relationshipPath))
      const relationshipIds = addRelationshipsWithFreshIds(relationships, sheetArtifacts.relationships)
      if (sheetArtifacts.relationships.length > 0) {
        setZipText(zip, relationshipPath, buildRelationshipsXml(relationships))
        changed = true
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (sheetXml && sheetArtifacts.pivotTableDefinitionsXml) {
        setZipText(zip, sheetPath, insertPivotTableDefinitionsXml(sheetXml, sheetArtifacts.pivotTableDefinitionsXml, relationshipIds))
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
