import { unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookPackageRelationshipSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSheetThreadedCommentArtifactsSnapshot,
  WorkbookSnapshot,
  WorkbookThreadedCommentArtifactsSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  escapeXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'

const binaryChunkSize = 0x8000
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const contentTypesPath = '[Content_Types].xml'
const threadedCommentRelationshipTypeFragment = '/relationships/threadedComment'
const personRelationshipTypeFragment = '/relationships/person'
const threadedCommentContentType = 'application/vnd.ms-excel.threadedcomments+xml'
const personContentType = 'application/vnd.ms-excel.person+xml'
const threadedCommentPartPathPattern = /^xl\/threadedComments\/threadedComment[^/]*\.xml$/u
const personPartPathPattern = /^xl\/persons\/person[^/]*\.xml$/u

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

export interface ImportedThreadedCommentArtifacts {
  readonly artifacts?: WorkbookThreadedCommentArtifactsSnapshot
  readonly sheetArtifactsByName: Map<string, WorkbookSheetThreadedCommentArtifactsSnapshot>
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

function decodedPartBytes(part: WorkbookPreservedPackagePartSnapshot): Uint8Array | undefined {
  if (part.storage !== 'base64') {
    return undefined
  }
  const bytes = decodeBase64(part.dataBase64)
  return bytes.byteLength === part.byteLength ? bytes : undefined
}

function readXmlAttribute(attributes: string, attributeName: string): string | null {
  return new RegExp(`\\b${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(attributes)?.[2] ?? null
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : ''
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

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function extensionFromPath(path: string): string | null {
  const fileName = normalizeZipPath(path).slice(normalizeZipPath(path).lastIndexOf('/') + 1)
  const extensionIndex = fileName.lastIndexOf('.')
  return extensionIndex >= 0 && extensionIndex < fileName.length - 1 ? fileName.slice(extensionIndex + 1).toLowerCase() : null
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const name = readXmlAttribute(attributes, 'name')
    const relationshipId = readXmlAttribute(attributes, 'r:id') ?? readXmlAttribute(attributes, 'id')
    return name && relationshipId ? [{ name: decodeXmlText(name), relationshipId }] : []
  })
}

function worksheetPathsBySheetName(zip: XlsxZipEntries, sheetNames: readonly string[]): Map<string, string> {
  const paths = new Map<string, string>()
  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath))
  const worksheetRelationshipsById = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]),
  )
  readWorkbookSheetEntries(getZipText(zip, workbookPath)).forEach((entry) => {
    const worksheetPath = worksheetRelationshipsById.get(entry.relationshipId)
    if (worksheetPath) {
      paths.set(entry.name, worksheetPath)
    }
  })
  sheetNames.forEach((sheetName, sheetIndex) => {
    if (!paths.has(sheetName)) {
      paths.set(sheetName, sheetPath(sheetIndex))
    }
  })
  return paths
}

function readContentTypeDefaults(contentTypesXml: string, partPaths: readonly string[]): WorkbookContentTypeDefaultSnapshot[] {
  const neededExtensions = new Set(partPaths.map(extensionFromPath).filter((extension): extension is string => Boolean(extension)))
  const defaultsByExtension = new Map<string, WorkbookContentTypeDefaultSnapshot>()
  for (const match of contentTypesXml.matchAll(/<Default\b([^>]*)\/?>/gu)) {
    const attributes = match[1] ?? ''
    const extension = readXmlAttribute(attributes, 'Extension')?.toLowerCase()
    const contentType = readXmlAttribute(attributes, 'ContentType')
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
    const partName = readXmlAttribute(attributes, 'PartName')
    const contentType = readXmlAttribute(attributes, 'ContentType')
    if (!partName || !contentType || !neededPartNames.has(partName)) {
      continue
    }
    overridesByPartName.set(partName, { partName, contentType })
  }
  return [...overridesByPartName.values()].toSorted((left, right) => left.partName.localeCompare(right.partName))
}

function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  const existingDefaultPattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${escapeRegExp(extension)}\\1`, 'u')
  if (existingDefaultPattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/></Types>`,
  )
}

function relationshipSnapshot(relationship: ParsedRelationship): WorkbookPackageRelationshipSnapshot {
  return {
    id: relationship.id,
    type: relationship.type,
    target: relationship.target,
    ...(relationship.targetMode ? { targetMode: relationship.targetMode } : {}),
  }
}

function sheetRelationshipPath(sheetIndex: number): string {
  return `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
}

function sheetPath(sheetIndex: number): string {
  return `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
}

function sheetRelationshipPathFromSheetPath(path: string): string {
  const normalizedPath = normalizeZipPath(path)
  const separatorIndex = normalizedPath.lastIndexOf('/')
  const directory = separatorIndex >= 0 ? normalizedPath.slice(0, separatorIndex + 1) : ''
  const fileName = separatorIndex >= 0 ? normalizedPath.slice(separatorIndex + 1) : normalizedPath
  return `${directory}_rels/${fileName}.rels`
}

function resolvesToThreadedCommentPart(basePartPath: string, relationship: ParsedRelationship): boolean {
  const partPath = normalizeZipPath(resolveTargetPath(basePartPath, relationship.target))
  return relationship.type.includes(threadedCommentRelationshipTypeFragment) || threadedCommentPartPathPattern.test(partPath)
}

function resolvesToPersonPart(basePartPath: string, relationship: ParsedRelationship): boolean {
  const partPath = normalizeZipPath(resolveTargetPath(basePartPath, relationship.target))
  return relationship.type.includes(personRelationshipTypeFragment) || personPartPathPattern.test(partPath)
}

function addRelationshipTargetPartPath(partPaths: Set<string>, basePartPath: string, relationship: ParsedRelationship): void {
  if (relationship.targetMode === 'External') {
    return
  }
  partPaths.add(normalizeZipPath(resolveTargetPath(basePartPath, relationship.target)))
}

function readPreservedParts(zip: XlsxZipEntries, partPaths: Set<string>): WorkbookPreservedPackagePartSnapshot[] {
  return [...partPaths]
    .filter((path) => threadedCommentPartPathPattern.test(path) || personPartPathPattern.test(path))
    .toSorted()
    .flatMap((path) => {
      const bytes = zip[path]
      return bytes ? [encodedPartSnapshot(path, bytes)] : []
    })
}

export function readImportedWorkbookThreadedCommentArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): ImportedThreadedCommentArtifacts {
  const zip = readXlsxZipEntries(source)
  const sheetArtifactsByName = new Map<string, WorkbookSheetThreadedCommentArtifactsSnapshot>()
  const partPaths = new Set<string>()

  const worksheetPaths = worksheetPathsBySheetName(zip, sheetNames)
  sheetNames.forEach((sheetName, sheetIndex) => {
    const currentSheetPath = worksheetPaths.get(sheetName) ?? sheetPath(sheetIndex)
    const relationships = parseRelationships(getZipText(zip, sheetRelationshipPathFromSheetPath(currentSheetPath))).filter((relationship) =>
      resolvesToThreadedCommentPart(currentSheetPath, relationship),
    )
    if (relationships.length === 0) {
      return
    }
    relationships.forEach((relationship) => addRelationshipTargetPartPath(partPaths, currentSheetPath, relationship))
    sheetArtifactsByName.set(sheetName, {
      relationships: relationships.map(relationshipSnapshot),
    })
  })

  const workbookRelationships = parseRelationships(getZipText(zip, workbookRelationshipsPath)).filter((relationship) =>
    resolvesToPersonPart(workbookPath, relationship),
  )
  workbookRelationships.forEach((relationship) => addRelationshipTargetPartPath(partPaths, workbookPath, relationship))

  Object.keys(zip).forEach((path) => {
    if (threadedCommentPartPathPattern.test(path) || personPartPathPattern.test(path)) {
      partPaths.add(path)
    }
  })

  const parts = readPreservedParts(zip, partPaths)
  if (parts.length === 0 && workbookRelationships.length === 0 && sheetArtifactsByName.size === 0) {
    return { sheetArtifactsByName }
  }
  const contentTypesXml = getZipText(zip, contentTypesPath)
  const partPathsArray = parts.map((part) => part.path)
  const artifacts: WorkbookThreadedCommentArtifactsSnapshot = {
    parts,
    ...(workbookRelationships.length > 0 ? { workbookRelationships: workbookRelationships.map(relationshipSnapshot) } : {}),
    ...(contentTypesXml ? { contentTypeDefaults: readContentTypeDefaults(contentTypesXml, partPathsArray) } : {}),
    ...(contentTypesXml ? { contentTypeOverrides: readContentTypeOverrides(contentTypesXml, partPathsArray) } : {}),
  }
  return { artifacts, sheetArtifactsByName }
}

function mergeRelationships(
  relationships: readonly ParsedRelationship[],
  additions: readonly WorkbookPackageRelationshipSnapshot[],
): ParsedRelationship[] {
  const output: ParsedRelationship[] = relationships.map((relationship) => ({ ...relationship }))
  const relationshipKeys = new Set(output.map((relationship) => `${relationship.type}\u0000${relationship.target}`))
  const usedIds = new Set(output.map((relationship) => relationship.id))
  additions.forEach((addition) => {
    const relationshipKey = `${addition.type}\u0000${addition.target}`
    if (relationshipKeys.has(relationshipKey)) {
      return
    }
    let id = addition.id
    while (usedIds.has(id)) {
      id = nextRelationshipId(output)
    }
    output.push({
      id,
      type: addition.type,
      target: addition.target,
      ...(addition.targetMode ? { targetMode: addition.targetMode } : {}),
    })
    usedIds.add(id)
    relationshipKeys.add(relationshipKey)
  })
  return output
}

function addFallbackContentTypeOverride(contentTypesXml: string, path: string): string {
  if (threadedCommentPartPathPattern.test(path)) {
    return addContentTypeOverride(contentTypesXml, `/${path}`, threadedCommentContentType)
  }
  if (personPartPathPattern.test(path)) {
    return addContentTypeOverride(contentTypesXml, `/${path}`, personContentType)
  }
  return contentTypesXml
}

function applyContentTypes(contentTypesXml: string, artifacts: WorkbookThreadedCommentArtifactsSnapshot): string {
  let output = contentTypesXml
  for (const entry of artifacts.contentTypeDefaults ?? []) {
    output = addContentTypeDefault(output, entry.extension, entry.contentType)
  }
  for (const entry of artifacts.contentTypeOverrides ?? []) {
    output = addContentTypeOverride(output, entry.partName, entry.contentType)
  }
  for (const part of artifacts.parts) {
    output = addFallbackContentTypeOverride(output, normalizeZipPath(part.path))
  }
  return output
}

export function addExportThreadedCommentArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookArtifacts = snapshot.workbook.metadata?.threadedCommentArtifacts
  const sheetsWithArtifacts = snapshot.sheets.filter((sheet) => sheet.metadata?.threadedCommentArtifacts)
  if (!workbookArtifacts && sheetsWithArtifacts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  if (workbookArtifacts) {
    for (const part of workbookArtifacts.parts) {
      const partBytes = decodedPartBytes(part)
      if (!partBytes) {
        continue
      }
      zip[normalizeZipPath(part.path)] = partBytes
      changed = true
    }
    if (workbookArtifacts.workbookRelationships && workbookArtifacts.workbookRelationships.length > 0) {
      const nextRelationships = mergeRelationships(
        parseRelationships(getZipText(zip, workbookRelationshipsPath)),
        workbookArtifacts.workbookRelationships,
      )
      setZipText(zip, workbookRelationshipsPath, buildRelationshipsXml(nextRelationships))
      changed = true
    }
    const contentTypesXml = getZipText(zip, contentTypesPath)
    if (contentTypesXml) {
      setZipText(zip, contentTypesPath, applyContentTypes(contentTypesXml, workbookArtifacts))
      changed = true
    }
  }

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const threadedCommentArtifacts = sheet.metadata?.threadedCommentArtifacts
      if (!threadedCommentArtifacts || threadedCommentArtifacts.relationships.length === 0) {
        return
      }
      const relsPath = sheetRelationshipPath(sheetIndex)
      const nextRelationships = mergeRelationships(parseRelationships(getZipText(zip, relsPath)), threadedCommentArtifacts.relationships)
      setZipText(zip, relsPath, buildRelationshipsXml(nextRelationships))
      changed = true
    })

  return changed ? zipSync(zip) : bytes
}
