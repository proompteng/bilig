import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSlicerConnectionArtifactsSnapshot,
  WorkbookSlicerConnectionSheetArtifactsSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import {
  getZipText,
  normalizeZipPath,
  readXlsxZipEntries,
  readXlsxZipEntryUncompressedSize,
  type XlsxZipEntries,
  type XlsxZipSource,
} from './xlsx-zip.js'
import {
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'

const binaryChunkSize = 0x8000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const contentTypesPath = '[Content_Types].xml'
const connectionsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections'
const slicerCacheRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const slicerRelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const connectionsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml'
const slicerCacheContentType = 'application/vnd.ms-excel.slicerCache+xml'
const slicerContentType = 'application/vnd.ms-excel.slicer+xml'
const extensionElementPattern = /<(?:[A-Za-z_][\w.-]*:)?ext\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?ext>)/gu
const extLstClosingElementPattern = /<\/(?:[A-Za-z_][\w.-]*:)?extLst>/u
const connectionsPartPathPattern = /^xl\/connections\.xml$/u
const connectionsRelationshipPartPathPattern = /^xl\/_rels\/connections\.xml\.rels$/u
const slicerCachePartPathPattern = /^xl\/slicerCaches\/slicerCache[1-9][0-9]*\.xml$/u
const slicerCacheRelationshipPartPathPattern = /^xl\/slicerCaches\/_rels\/slicerCache[1-9][0-9]*\.xml\.rels$/u
const slicerPartPathPattern = /^xl\/slicers\/slicer[1-9][0-9]*\.xml$/u
const slicerRelationshipPartPathPattern = /^xl\/slicers\/_rels\/slicer[1-9][0-9]*\.xml\.rels$/u

export interface ImportedWorkbookSlicerConnectionSheetSource {
  readonly sheetName: string
  readonly sheetPath: string
  readonly sheetSlicerListExtXml?: string
}

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
  }
  return binary
}

function decodeBinaryString(binary: string): Uint8Array {
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index)
  }
  return bytes
}

function encodeBase64(bytes: Uint8Array): string {
  const btoa = globalThis.btoa
  if (typeof btoa === 'function') {
    return btoa(encodeBinaryString(bytes))
  }
  return Buffer.from(bytes).toString('base64')
}

function decodeBase64(dataBase64: string): Uint8Array {
  const atob = globalThis.atob
  if (typeof atob === 'function') {
    return decodeBinaryString(atob(dataBase64))
  }
  return new Uint8Array(Buffer.from(dataBase64, 'base64'))
}

function encodedPartSnapshot(path: string, bytes: Uint8Array): WorkbookPreservedPackagePartSnapshot {
  return {
    path,
    storage: 'base64',
    dataBase64: encodeBase64(bytes),
    byteLength: bytes.byteLength,
  }
}

class LazyEncodedPartSnapshot implements WorkbookPreservedPackagePartSnapshot {
  readonly storage = 'base64' as const
  declare readonly dataBase64: string
  private dataBase64Cache: string | undefined

  constructor(
    readonly path: string,
    readonly byteLength: number,
    private readonly readBytes: () => Uint8Array | undefined,
  ) {
    Object.defineProperty(this, 'dataBase64', {
      configurable: true,
      enumerable: true,
      get: () => this.getDataBase64(),
    })
  }

  private getDataBase64(): string {
    this.dataBase64Cache ??= encodeBase64(this.readBytes() ?? new Uint8Array())
    return this.dataBase64Cache
  }
}

function lazyEncodedPartSnapshot(
  path: string,
  byteLength: number,
  readBytes: () => Uint8Array | undefined,
): WorkbookPreservedPackagePartSnapshot {
  return new LazyEncodedPartSnapshot(path, byteLength, readBytes)
}

function decodedPartBytes(part: WorkbookPreservedPackagePartSnapshot): Uint8Array | undefined {
  if (part.storage !== 'base64') {
    return undefined
  }
  const bytes = decodeBase64(part.dataBase64)
  return bytes.byteLength === part.byteLength ? bytes : undefined
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function extensionFromPath(path: string): string | null {
  const normalized = normalizeZipPath(path)
  const fileName = normalized.slice(normalized.lastIndexOf('/') + 1)
  const extensionIndex = fileName.lastIndexOf('.')
  return extensionIndex >= 0 && extensionIndex < fileName.length - 1 ? fileName.slice(extensionIndex + 1).toLowerCase() : null
}

function readContentTypeDefaults(contentTypesXml: string, partPaths: readonly string[]): WorkbookContentTypeDefaultSnapshot[] {
  const neededExtensions = new Set(partPaths.map(extensionFromPath).filter((extension): extension is string => Boolean(extension)))
  const defaultsByExtension = new Map<string, WorkbookContentTypeDefaultSnapshot>()
  for (const match of contentTypesXml.matchAll(/<Default\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    const extension = readAttribute(attributes, 'Extension')?.toLowerCase()
    const contentType = readAttribute(attributes, 'ContentType')
    if (!extension || !contentType || !neededExtensions.has(extension)) {
      continue
    }
    defaultsByExtension.set(extension, { extension, contentType })
  }
  return [...defaultsByExtension.values()].toSorted((left, right) => left.extension.localeCompare(right.extension))
}

function readContentTypeOverrides(contentTypesXml: string, partPaths: readonly string[]): WorkbookContentTypeOverrideSnapshot[] {
  const neededPartNames = new Set(partPaths.map((path) => `/${normalizeZipPath(path)}`))
  const overridesByPartName = new Map<string, WorkbookContentTypeOverrideSnapshot>()
  for (const match of contentTypesXml.matchAll(/<Override\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    const partName = readAttribute(attributes, 'PartName')
    const contentType = readAttribute(attributes, 'ContentType')
    if (!partName || !contentType || !neededPartNames.has(partName)) {
      continue
    }
    overridesByPartName.set(partName, { partName, contentType })
  }
  return [...overridesByPartName.values()].toSorted((left, right) => left.partName.localeCompare(right.partName))
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  const pattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${escapeRegExp(extension)}\\1`, 'u')
  if (pattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/></Types>`,
  )
}

function upsertContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (!contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  const escapedPartName = escapeXml(partName)
  const escapedContentType = escapeXml(contentType)
  const overridePattern = /<Override\b([^>]*)\/?>/gu
  let replaced = false
  const nextXml = contentTypesXml.replace(overridePattern, (match: string, attributes: string) => {
    if (readAttribute(attributes, 'PartName') !== partName) {
      return match
    }
    replaced = true
    return `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/>`
  })
  return replaced
    ? nextXml
    : nextXml.replace('</Types>', `<Override PartName="${escapedPartName}" ContentType="${escapedContentType}"/></Types>`)
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

function isWorkbookArtifactRelationship(relationship: ParsedRelationship | WorkbookPackageRelationshipSnapshot): boolean {
  return relationship.type === connectionsRelationshipType || relationship.type === slicerCacheRelationshipType
}

function isSlicerRelationship(relationship: ParsedRelationship | WorkbookPackageRelationshipSnapshot): boolean {
  return relationship.type === slicerRelationshipType
}

function relationshipPartPath(partPath: string): string {
  const normalized = normalizeZipPath(partPath)
  const slashIndex = normalized.lastIndexOf('/')
  const directory = slashIndex >= 0 ? normalized.slice(0, slashIndex) : ''
  const fileName = slashIndex >= 0 ? normalized.slice(slashIndex + 1) : normalized
  return directory.length > 0 ? `${directory}/_rels/${fileName}.rels` : `_rels/${fileName}.rels`
}

function hasZipEntry(zip: XlsxZipEntries, path: string): boolean {
  return Object.hasOwn(zip, normalizeZipPath(path))
}

function isPreservedPackagePartPath(path: string): boolean {
  const normalized = normalizeZipPath(path)
  return (
    connectionsPartPathPattern.test(normalized) ||
    connectionsRelationshipPartPathPattern.test(normalized) ||
    slicerCachePartPathPattern.test(normalized) ||
    slicerCacheRelationshipPartPathPattern.test(normalized) ||
    slicerPartPathPattern.test(normalized) ||
    slicerRelationshipPartPathPattern.test(normalized)
  )
}

function fallbackContentTypeForPath(path: string): string | null {
  const normalized = normalizeZipPath(path)
  if (connectionsPartPathPattern.test(normalized)) {
    return connectionsContentType
  }
  if (slicerCachePartPathPattern.test(normalized)) {
    return slicerCacheContentType
  }
  if (slicerPartPathPattern.test(normalized)) {
    return slicerContentType
  }
  return null
}

function extensionXmlWithChild(xml: string | null, childName: 'slicerCaches' | 'slicerList'): string | undefined {
  if (!xml) {
    return undefined
  }
  const childPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${childName}\\b`, 'u')
  extensionElementPattern.lastIndex = 0
  return [...xml.matchAll(extensionElementPattern)].find((match) => childPattern.test(match[0]))?.[0]
}

function removeExtensionXmlWithChild(xml: string, childName: 'slicerCaches' | 'slicerList'): string {
  const childPattern = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${childName}\\b`, 'u')
  extensionElementPattern.lastIndex = 0
  return xml.replace(extensionElementPattern, (extensionXml) => (childPattern.test(extensionXml) ? '' : extensionXml))
}

function insertExtensionXml(
  xml: string,
  extensionXml: string,
  relationshipIds: ReadonlyMap<string, string>,
  childName: 'slicerCaches' | 'slicerList',
  rootElementName: 'workbook' | 'worksheet',
): string {
  const nextExtensionXml = replaceRelationshipIds(extensionXml, relationshipIds)
  const withoutExistingExtension = ensureRelationshipNamespace(removeExtensionXmlWithChild(xml, childName))
  if (extLstClosingElementPattern.test(withoutExistingExtension)) {
    return withoutExistingExtension.replace(extLstClosingElementPattern, `${nextExtensionXml}$&`)
  }
  return withoutExistingExtension.replace(
    new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${rootElementName}>`, 'u'),
    `<extLst>${nextExtensionXml}</extLst>$&`,
  )
}

function readSheetArtifacts(
  zip: XlsxZipEntries,
  sheets: readonly ImportedWorkbookSlicerConnectionSheetSource[],
): WorkbookSlicerConnectionSheetArtifactsSnapshot[] {
  return sheets.flatMap((sheet) => {
    const relationshipsPath = relationshipPartPath(sheet.sheetPath)
    const sheetSlicerListExtXml = sheet.sheetSlicerListExtXml ?? extensionXmlWithChild(getZipText(zip, sheet.sheetPath), 'slicerList')
    const relationships = parseRelationships(getZipText(zip, relationshipsPath)).filter(isSlicerRelationship).map(relationshipSnapshot)
    if (!sheetSlicerListExtXml && relationships.length === 0) {
      return []
    }
    return [
      {
        sheetName: sheet.sheetName,
        ...(sheetSlicerListExtXml ? { sheetSlicerListExtXml } : {}),
        ...(relationships.length > 0 ? { relationships } : {}),
      },
    ]
  })
}

function addRelationshipTargetPartPath(
  partPaths: Set<string>,
  zip: XlsxZipEntries,
  basePartPath: string,
  relationship: WorkbookPackageRelationshipSnapshot,
): void {
  const targetPath = normalizeZipPath(resolveTargetPath(basePartPath, relationship.target))
  if (isPreservedPackagePartPath(targetPath) && hasZipEntry(zip, targetPath)) {
    partPaths.add(targetPath)
  }
}

function readSlicerConnectionPartPaths(
  zip: XlsxZipEntries,
  sheets: readonly ImportedWorkbookSlicerConnectionSheetSource[],
  workbookRelationships: readonly WorkbookPackageRelationshipSnapshot[],
  sheetArtifacts: readonly WorkbookSlicerConnectionSheetArtifactsSnapshot[],
): string[] {
  const sheetPathsByName = new Map(sheets.map((sheet) => [sheet.sheetName, sheet.sheetPath]))
  const partPaths = new Set<string>()
  for (const relationship of workbookRelationships) {
    addRelationshipTargetPartPath(partPaths, zip, workbookPath, relationship)
  }
  for (const sheetArtifact of sheetArtifacts) {
    const sheetPath = sheetPathsByName.get(sheetArtifact.sheetName)
    if (!sheetPath) {
      continue
    }
    for (const relationship of sheetArtifact.relationships ?? []) {
      addRelationshipTargetPartPath(partPaths, zip, sheetPath, relationship)
    }
  }
  for (const path of Object.keys(zip)) {
    if (isPreservedPackagePartPath(path) && !path.includes('/_rels/')) {
      partPaths.add(path)
    }
  }

  const allPartPaths = new Set<string>(partPaths)
  for (const partPath of partPaths) {
    const relsPath = relationshipPartPath(partPath)
    if (hasZipEntry(zip, relsPath)) {
      allPartPaths.add(relsPath)
    }
  }
  return [...allPartPaths].toSorted()
}

function preservedPartsByPath(parts: readonly WorkbookPreservedPackagePartSnapshot[]): Map<string, Uint8Array> {
  const output = new Map<string, Uint8Array>()
  for (const part of parts) {
    if (!isPreservedPackagePartPath(part.path)) {
      continue
    }
    const bytes = decodedPartBytes(part)
    if (bytes) {
      output.set(normalizeZipPath(part.path), bytes)
    }
  }
  return output
}

function relationshipTargetExists(
  relationship: WorkbookPackageRelationshipSnapshot,
  partsByPath: ReadonlyMap<string, Uint8Array>,
  zip: XlsxZipEntries,
  basePartPath: string,
): boolean {
  const targetPath = normalizeZipPath(resolveTargetPath(basePartPath, relationship.target))
  return isPreservedPackagePartPath(targetPath) && (partsByPath.has(targetPath) || Boolean(zip[targetPath]))
}

function addRelationshipsWithStableTargets(input: {
  readonly relationships: ParsedRelationship[]
  readonly additions: readonly WorkbookPackageRelationshipSnapshot[] | undefined
  readonly partsByPath: ReadonlyMap<string, Uint8Array>
  readonly zip: XlsxZipEntries
  readonly basePartPath: string
  readonly includeRelationship: (relationship: ParsedRelationship | WorkbookPackageRelationshipSnapshot) => boolean
}): { readonly changed: boolean; readonly idMap: Map<string, string> } {
  const idMap = new Map<string, string>()
  let changed = false
  for (const addition of input.additions ?? []) {
    if (!input.includeRelationship(addition) || !relationshipTargetExists(addition, input.partsByPath, input.zip, input.basePartPath)) {
      continue
    }
    const existing = input.relationships.find(
      (relationship) => relationship.type === addition.type && relationship.target === addition.target,
    )
    if (existing) {
      idMap.set(addition.id, existing.id)
      continue
    }
    const idInUse = input.relationships.some((relationship) => relationship.id === addition.id)
    const nextId = addition.id.length > 0 && !idInUse ? addition.id : nextRelationshipId(input.relationships)
    input.relationships.push({ ...parsedRelationship(addition), id: nextId })
    idMap.set(addition.id, nextId)
    changed = true
  }
  return { changed, idMap }
}

function replaceRelationshipIds(xml: string, relationshipIds: ReadonlyMap<string, string>): string {
  if (relationshipIds.size === 0) {
    return xml
  }
  return xml.replace(/\br:id=(["'])([\s\S]*?)\1/gu, (match, quote: string, id: string) => {
    const nextId = relationshipIds.get(id)
    return nextId ? match.replace(`${quote}${id}${quote}`, `${quote}${nextId}${quote}`) : match
  })
}

function addSlicerConnectionContentTypes(
  contentTypesXml: string,
  artifacts: WorkbookSlicerConnectionArtifactsSnapshot,
  copiedPartPaths: ReadonlySet<string>,
): string {
  let output = contentTypesXml
  const copiedExtensions = new Set(
    [...copiedPartPaths].map(extensionFromPath).filter((extension): extension is string => Boolean(extension)),
  )
  for (const defaultEntry of artifacts.contentTypeDefaults ?? []) {
    if (copiedExtensions.has(defaultEntry.extension)) {
      output = addContentTypeDefault(output, defaultEntry.extension, defaultEntry.contentType)
    }
  }
  for (const overrideEntry of artifacts.contentTypeOverrides ?? []) {
    const path = normalizeZipPath(overrideEntry.partName)
    if (copiedPartPaths.has(path)) {
      output = upsertContentTypeOverride(output, overrideEntry.partName, overrideEntry.contentType)
    }
  }
  for (const path of copiedPartPaths) {
    const contentType = fallbackContentTypeForPath(path)
    if (contentType) {
      output = upsertContentTypeOverride(output, `/${path}`, contentType)
    }
  }
  return output
}

export function readImportedWorkbookSlicerConnectionArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): WorkbookSlicerConnectionArtifactsSnapshot | undefined {
  return readImportedWorkbookSlicerConnectionArtifactsFromSheets(
    source,
    sheetNames.map((sheetName, sheetIndex) => ({
      sheetName,
      sheetPath: `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`,
    })),
  )
}

export function readImportedWorkbookSlicerConnectionArtifactsFromSheets(
  source: XlsxZipSource,
  sheets: readonly ImportedWorkbookSlicerConnectionSheetSource[],
  options: {
    readonly workbookXml?: string
    readonly workbookRelationshipsXml?: string
  } = {},
): WorkbookSlicerConnectionArtifactsSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const workbookSlicerCachesExtXml = extensionXmlWithChild(options.workbookXml ?? getZipText(zip, workbookPath), 'slicerCaches')
  const workbookRelationships = parseRelationships(options.workbookRelationshipsXml ?? getZipText(zip, workbookRelationshipsPath))
    .filter(isWorkbookArtifactRelationship)
    .map(relationshipSnapshot)
  const sheetArtifacts = readSheetArtifacts(zip, sheets)
  const partPaths = readSlicerConnectionPartPaths(zip, sheets, workbookRelationships, sheetArtifacts)
  const parts = partPaths.flatMap((path) => {
    const byteLength = readXlsxZipEntryUncompressedSize(zip, path)
    if (byteLength !== undefined) {
      return [
        lazyEncodedPartSnapshot(path, byteLength, () => {
          const bytes = zip[path]
          if (bytes) {
            Reflect.deleteProperty(zip, path)
          }
          return bytes
        }),
      ]
    }
    const bytes = zip[path]
    if (!bytes) {
      return []
    }
    const snapshot = encodedPartSnapshot(path, bytes)
    Reflect.deleteProperty(zip, path)
    return [snapshot]
  })
  if (parts.length === 0 && !workbookSlicerCachesExtXml && workbookRelationships.length === 0 && sheetArtifacts.length === 0) {
    return undefined
  }

  const contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const contentTypeDefaults = contentTypesXml ? readContentTypeDefaults(contentTypesXml, partPaths) : []
  const contentTypeOverrides = contentTypesXml ? readContentTypeOverrides(contentTypesXml, partPaths) : []
  return {
    parts,
    ...(workbookSlicerCachesExtXml ? { workbookSlicerCachesExtXml } : {}),
    ...(workbookRelationships.length > 0 ? { workbookRelationships } : {}),
    ...(sheetArtifacts.length > 0 ? { sheetArtifacts } : {}),
    ...(contentTypeDefaults.length > 0 ? { contentTypeDefaults } : {}),
    ...(contentTypeOverrides.length > 0 ? { contentTypeOverrides } : {}),
  }
}

export function addExportSlicerConnectionArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const artifacts = snapshot.workbook.metadata?.slicerConnectionArtifacts
  if (!artifacts || artifacts.parts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(artifacts.parts)
  if (partsByPath.size === 0) {
    return bytes
  }

  let changed = false
  const copiedPartPaths = new Set<string>()
  for (const [path, partBytes] of partsByPath) {
    zip[path] = partBytes
    copiedPartPaths.add(path)
    changed = true
  }

  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const workbookRelationshipResult = addRelationshipsWithStableTargets({
    relationships: workbookRelationships,
    additions: artifacts.workbookRelationships,
    partsByPath,
    zip,
    basePartPath: workbookPath,
    includeRelationship: isWorkbookArtifactRelationship,
  })
  if (workbookRelationshipResult.changed) {
    setZipText(zip, workbookRelationshipsPath, buildRelationshipsXml(workbookRelationships))
    changed = true
  }

  const workbookXml = getZipText(zip, workbookPath)
  if (workbookXml && artifacts.workbookSlicerCachesExtXml) {
    setZipText(
      zip,
      workbookPath,
      insertExtensionXml(workbookXml, artifacts.workbookSlicerCachesExtXml, workbookRelationshipResult.idMap, 'slicerCaches', 'workbook'),
    )
    changed = true
  }

  const sheetArtifactsByName = new Map((artifacts.sheetArtifacts ?? []).map((sheetArtifact) => [sheetArtifact.sheetName, sheetArtifact]))
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetArtifacts = sheetArtifactsByName.get(sheet.name)
      if (!sheetArtifacts) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const relationshipPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const relationships = parseRelationships(getZipText(zip, relationshipPath))
      const relationshipResult = addRelationshipsWithStableTargets({
        relationships,
        additions: sheetArtifacts.relationships,
        partsByPath,
        zip,
        basePartPath: sheetPath,
        includeRelationship: isSlicerRelationship,
      })
      if (relationshipResult.changed) {
        setZipText(zip, relationshipPath, buildRelationshipsXml(relationships))
        changed = true
      }

      const sheetXml = getZipText(zip, sheetPath)
      if (sheetXml && sheetArtifacts.sheetSlicerListExtXml) {
        setZipText(
          zip,
          sheetPath,
          insertExtensionXml(sheetXml, sheetArtifacts.sheetSlicerListExtXml, relationshipResult.idMap, 'slicerList', 'worksheet'),
        )
        changed = true
      }
    })

  const contentTypesXml = getZipText(zip, contentTypesPath) ?? ''
  const nextContentTypesXml = contentTypesXml
    ? addSlicerConnectionContentTypes(contentTypesXml, artifacts, copiedPartPaths)
    : contentTypesXml
  if (nextContentTypesXml !== contentTypesXml) {
    setZipText(zip, contentTypesPath, nextContentTypesXml)
    changed = true
  }

  return changed ? zipSync(zip) : bytes
}
