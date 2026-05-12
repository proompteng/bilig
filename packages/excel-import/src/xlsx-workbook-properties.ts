import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { XMLParser } from 'fast-xml-parser'

import type {
  LiteralInput,
  WorkbookDocumentPropertiesArtifactsSnapshot,
  WorkbookDocumentPropertyPartSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPropertySnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>
type DocumentPropertyPartWithContentType = WorkbookDocumentPropertyPartSnapshot & { readonly contentType: string }

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
  readonly targetMode?: string
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
})

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const corePropertiesRelationshipType = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const extendedPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
const customPropertiesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
const documentPropertiesValueTypesNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
const customPropertiesRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
const corePropertiesContentType = 'application/vnd.openxmlformats-package.core-properties+xml'
const extendedPropertiesContentType = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const customPropertiesContentType = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
const rootRelationshipsPath = '_rels/.rels'
const contentTypesPath = '[Content_Types].xml'
const corePropertiesPartPath = 'docProps/core.xml'
const extendedPropertiesPartPath = 'docProps/app.xml'
const customPropertiesPartPath = 'docProps/custom.xml'
const customPropertiesFormatId = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'

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

function textValue(value: unknown): string | null {
  if (typeof value === 'string') {
    return value
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  return null
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

function parseRelationships(xml: string | null): ParsedRelationship[] {
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
        )}"${relationship.targetMode ? ` TargetMode="${escapeXml(relationship.targetMode)}"` : ''}/>`,
    ),
    '</Relationships>',
  ].join('')
}

function readXmlAttribute(attributes: string, name: string): string | null {
  return new RegExp(`\\b${name}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function buildContentTypesXml(partName: string, contentType: string): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    `<Override PartName="${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/>`,
    '</Types>',
  ].join('')
}

function upsertContentTypeOverride(contentTypesXml: string | null, partName: string, contentType: string): string {
  if (!contentTypesXml || !contentTypesXml.includes('</Types>')) {
    return buildContentTypesXml(partName, contentType)
  }
  const escapedPartName = escapeXml(partName)
  const escapedContentType = escapeXml(contentType)
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  let replaced = false
  const nextXml = contentTypesXml.replace(overridePattern, (match: string, attributes: string) => {
    if (readXmlAttribute(attributes, 'PartName') !== partName) {
      return match
    }
    replaced = true
    return `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/>`
  })
  return replaced
    ? nextXml
    : nextXml.replace('</Types>', `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/></Types>`)
}

function contentTypeOverride(contentTypesXml: string | null, partName: string): string | undefined {
  if (!contentTypesXml) {
    return undefined
  }
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    if (readXmlAttribute(attributes, 'PartName') === partName) {
      return readXmlAttribute(attributes, 'ContentType') ?? undefined
    }
  }
  return undefined
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

function readDocumentPropertyPart(input: {
  readonly zip: ZipEntries
  readonly relationships: readonly ParsedRelationship[]
  readonly relationshipType: string
  readonly fallbackPath: string
  readonly contentType: string
}): WorkbookDocumentPropertyPartSnapshot | undefined {
  const relationship = input.relationships.find((entry) => entry.type === input.relationshipType)
  const path = normalizeZipPath(relationship?.target ?? input.fallbackPath)
  const xml = getZipText(input.zip, path)
  if (!xml) {
    return undefined
  }
  return {
    path,
    xml,
    relationship: relationshipSnapshot(
      relationship ?? {
        id: '',
        target: input.fallbackPath,
        type: input.relationshipType,
      },
    ),
    contentType: contentTypeOverride(getZipText(input.zip, contentTypesPath), `/${path}`) ?? input.contentType,
  }
}

function relationshipIdForPart(relationships: readonly ParsedRelationship[], part: WorkbookDocumentPropertyPartSnapshot): string {
  if (part.relationship.id.length === 0) {
    return nextRelationshipId(relationships)
  }
  return relationships.some((relationship) => relationship.id === part.relationship.id)
    ? nextRelationshipId(relationships)
    : part.relationship.id
}

function upsertDocumentPropertyRelationship(relationships: ParsedRelationship[], part: WorkbookDocumentPropertyPartSnapshot): void {
  const nextRelationship = parsedRelationship(part.relationship)
  const existingIndex = relationships.findIndex((relationship) => relationship.type === nextRelationship.type)
  relationships[existingIndex >= 0 ? existingIndex : relationships.length] = {
    ...nextRelationship,
    id: existingIndex >= 0 ? (relationships[existingIndex]?.id ?? nextRelationship.id) : relationshipIdForPart(relationships, part),
  }
}

function addDocumentPropertyPartToZip(
  zip: ZipEntries,
  contentTypesXml: string | null,
  part: DocumentPropertyPartWithContentType | undefined,
): string | null {
  if (!part) {
    return contentTypesXml
  }
  const path = normalizeZipPath(part.path)
  setZipText(zip, path, part.xml)
  return upsertContentTypeOverride(contentTypesXml, `/${path}`, part.contentType)
}

function nonEmptyContentType(
  part: WorkbookDocumentPropertyPartSnapshot | undefined,
  fallback: string,
): DocumentPropertyPartWithContentType | undefined {
  if (!part) {
    return undefined
  }
  return {
    ...part,
    contentType: part.contentType && part.contentType.length > 0 ? part.contentType : fallback,
  }
}

function normalizableProperty(property: WorkbookPropertySnapshot): WorkbookPropertySnapshot | null {
  if (property.key.trim().length === 0) {
    return null
  }
  if (property.value === null) {
    return null
  }
  if (typeof property.value === 'number' && !Number.isFinite(property.value)) {
    return null
  }
  return { key: property.key, value: property.value }
}

function normalizeWorkbookProperties(properties: readonly WorkbookPropertySnapshot[] | undefined): WorkbookPropertySnapshot[] {
  const byKey = new Map<string, WorkbookPropertySnapshot>()
  for (const property of properties ?? []) {
    const normalized = normalizableProperty(property)
    if (normalized) {
      byKey.set(normalized.key, normalized)
    }
  }
  return [...byKey.values()].toSorted((left, right) => left.key.localeCompare(right.key))
}

function buildCustomPropertyValueXml(value: LiteralInput): string | null {
  if (typeof value === 'string') {
    return `<vt:lpwstr>${escapeXml(value)}</vt:lpwstr>`
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `<vt:r8>${escapeXml(String(value))}</vt:r8>`
  }
  if (typeof value === 'boolean') {
    return `<vt:bool>${value ? 'true' : 'false'}</vt:bool>`
  }
  return null
}

function buildCustomPropertiesXml(properties: readonly WorkbookPropertySnapshot[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Properties xmlns="${customPropertiesNamespace}" xmlns:vt="${documentPropertiesValueTypesNamespace}">`,
    ...properties.flatMap((property, index) => {
      const valueXml = buildCustomPropertyValueXml(property.value)
      if (!valueXml) {
        return []
      }
      return [
        `<property fmtid="${customPropertiesFormatId}" pid="${String(index + 2)}" name="${escapeXml(property.key)}">${valueXml}</property>`,
      ]
    }),
    '</Properties>',
  ].join('')
}

function ensureCustomPropertiesRelationship(zip: ZipEntries): void {
  const relationships = parseRelationships(getZipText(zip, rootRelationshipsPath))
  if (!relationships.some((relationship) => relationship.type === customPropertiesRelationshipType)) {
    relationships.push({
      id: nextRelationshipId(relationships),
      target: customPropertiesPartPath,
      type: customPropertiesRelationshipType,
    })
  }
  setZipText(zip, rootRelationshipsPath, buildRelationshipsXml(relationships))
}

function readCustomPropertiesPartPath(zip: ZipEntries): string | null {
  const relationship = parseRelationships(getZipText(zip, rootRelationshipsPath)).find(
    (entry) => entry.type === customPropertiesRelationshipType,
  )
  if (relationship) {
    return normalizeZipPath(relationship.target)
  }
  return zip[customPropertiesPartPath] ? customPropertiesPartPath : null
}

function readPropertyValue(property: Record<string, unknown>): LiteralInput | undefined {
  const stringValue = textValue(property['lpwstr'] ?? property['lpstr'])
  if (stringValue !== null) {
    return stringValue
  }

  const numberValue = textValue(property['r8'] ?? property['decimal'] ?? property['i4'] ?? property['int'] ?? property['ui4'])
  if (numberValue !== null) {
    const parsed = Number(numberValue)
    return Number.isFinite(parsed) ? parsed : undefined
  }

  const booleanValue = textValue(property['bool'])
  if (booleanValue !== null) {
    const normalized = booleanValue.trim().toLowerCase()
    if (normalized === 'true' || normalized === '1') {
      return true
    }
    if (normalized === 'false' || normalized === '0') {
      return false
    }
  }
  return undefined
}

export function addExportWorkbookPropertiesToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const properties = normalizeWorkbookProperties(snapshot.workbook.metadata?.properties)
  const documentPropertyArtifacts = snapshot.workbook.metadata?.documentPropertyArtifacts
  if (properties.length === 0 && !documentPropertyArtifacts) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let contentTypesXml = getZipText(zip, contentTypesPath)

  if (properties.length > 0) {
    setZipText(zip, customPropertiesPartPath, buildCustomPropertiesXml(properties))
    contentTypesXml = upsertContentTypeOverride(contentTypesXml, `/${customPropertiesPartPath}`, customPropertiesContentType)
    ensureCustomPropertiesRelationship(zip)
  }

  const documentPropertyParts = [
    nonEmptyContentType(documentPropertyArtifacts?.core, corePropertiesContentType),
    nonEmptyContentType(documentPropertyArtifacts?.app, extendedPropertiesContentType),
    nonEmptyContentType(documentPropertyArtifacts?.custom, customPropertiesContentType),
  ]
  const relationships = parseRelationships(getZipText(zip, rootRelationshipsPath))
  for (const part of documentPropertyParts) {
    contentTypesXml = addDocumentPropertyPartToZip(zip, contentTypesXml, part)
    if (part) {
      upsertDocumentPropertyRelationship(relationships, part)
    }
  }
  if (documentPropertyParts.some(Boolean)) {
    setZipText(zip, rootRelationshipsPath, buildRelationshipsXml(relationships))
  }
  if (contentTypesXml) {
    setZipText(zip, contentTypesPath, contentTypesXml)
  }
  return zipSync(zip)
}

export function readImportedWorkbookDocumentPropertiesArtifacts(
  source: XlsxZipSource,
): WorkbookDocumentPropertiesArtifactsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const relationships = parseRelationships(getZipText(zip, rootRelationshipsPath))
  const core = readDocumentPropertyPart({
    zip,
    relationships,
    relationshipType: corePropertiesRelationshipType,
    fallbackPath: corePropertiesPartPath,
    contentType: corePropertiesContentType,
  })
  const app = readDocumentPropertyPart({
    zip,
    relationships,
    relationshipType: extendedPropertiesRelationshipType,
    fallbackPath: extendedPropertiesPartPath,
    contentType: extendedPropertiesContentType,
  })
  const custom = readDocumentPropertyPart({
    zip,
    relationships,
    relationshipType: customPropertiesRelationshipType,
    fallbackPath: customPropertiesPartPath,
    contentType: customPropertiesContentType,
  })
  const artifacts: WorkbookDocumentPropertiesArtifactsSnapshot = {
    ...(core ? { core } : {}),
    ...(app ? { app } : {}),
    ...(custom ? { custom } : {}),
  }
  return Object.keys(artifacts).length > 0 ? artifacts : undefined
}

export function readImportedWorkbookProperties(source: XlsxZipSource): WorkbookPropertySnapshot[] | undefined {
  const zip = readXlsxZipEntries(source)
  const partPath = readCustomPropertiesPartPath(zip)
  if (!partPath) {
    return undefined
  }
  const customPropertiesXml = getZipText(zip, partPath)
  if (!customPropertiesXml) {
    return undefined
  }

  const parsed: unknown = xmlParser.parse(customPropertiesXml)
  const properties = asArray(recordChild(parsed, 'Properties')?.['property']).flatMap((entry) => {
    if (!isRecord(entry) || typeof entry['name'] !== 'string' || entry['name'].trim().length === 0) {
      return []
    }
    const value = readPropertyValue(entry)
    if (value === undefined) {
      return []
    }
    return [{ key: entry['name'], value }]
  })
  const normalizedProperties = normalizeWorkbookProperties(properties)
  return normalizedProperties.length > 0 ? normalizedProperties : undefined
}
