import { strFromU8, unzipSync, zipSync } from 'fflate'

import type {
  WorkbookContentTypeDefaultSnapshot,
  WorkbookContentTypeOverrideSnapshot,
  WorkbookDrawingArtifactsSnapshot,
  WorkbookPreservedPackagePartSnapshot,
  WorkbookSheetDrawingArtifactsSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  ensureRelationshipNamespace,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
  type ParsedRelationship,
} from './xlsx-pivot-artifacts.js'

const binaryChunkSize = 0x8000
const worksheetDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const chartRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const drawingContentType = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const drawingRootOpenTagPattern = /<(?<tag>(?:[A-Za-z_][\w.-]*:)?wsDr)\b[^>]*>/u
const drawingAnchorElementPattern =
  /<(?<tag>(?:[A-Za-z_][\w.-]*:)?(?:twoCellAnchor|oneCellAnchor|absoluteAnchor))\b[^>]*(?:\/>|>[\s\S]*?<\/\k<tag>>)/gu
const namespaceDeclarationPattern = /\s(xmlns(?::[A-Za-z_][\w.-]*)?)=("|')([\s\S]*?)\2/gu
const nonVisualPropertyIdPattern = /(<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\bid=")([0-9]+)(")/u
const drawingRelationshipAttributePattern = /\b(?:r:(?:id|embed|link)|id)=("|')([\s\S]*?)\1/gu

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

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function drawingRelationshipsPath(partPath: string): string {
  const normalizedPath = normalizeZipPath(partPath)
  const directory = normalizedPath.slice(0, normalizedPath.lastIndexOf('/'))
  const fileName = normalizedPath.slice(normalizedPath.lastIndexOf('/') + 1)
  return `${directory}/_rels/${fileName}.rels`
}

function readWorksheetDrawingRelationshipId(sheetXml: string | null): string | null {
  const drawingTag = /<drawing\b[^>]*(?:\/?>|>[\s\S]*?<\/drawing>)/u.exec(sheetXml ?? '')?.[0]
  return drawingTag ? (readAttribute(drawingTag, 'r:id') ?? readAttribute(drawingTag, 'id')) : null
}

function drawingRootOpenTag(drawingXml: string): string | null {
  return drawingRootOpenTagPattern.exec(drawingXml)?.[0] ?? null
}

function drawingRootTagName(openTag: string): string {
  return /<([^\s>]+)/u.exec(openTag)?.[1] ?? 'xdr:wsDr'
}

function drawingXmlDeclaration(xml: string): string | undefined {
  return /^\s*<\?xml[\s\S]*?\?>/u.exec(xml)?.[0]
}

function drawingAnchors(drawingXml: string): string[] {
  drawingAnchorElementPattern.lastIndex = 0
  return [...drawingXml.matchAll(drawingAnchorElementPattern)].map((match) => match[0])
}

function anchorContainsChart(anchorXml: string): boolean {
  return /<(?:[A-Za-z_][\w.-]*:)?chart\b/u.test(anchorXml)
}

function mergeNamespaceDeclarations(targetRootOpenTag: string, sourceRootOpenTag: string | null): string {
  if (!sourceRootOpenTag) {
    return targetRootOpenTag
  }
  let output = targetRootOpenTag
  namespaceDeclarationPattern.lastIndex = 0
  for (const match of sourceRootOpenTag.matchAll(namespaceDeclarationPattern)) {
    const declarationName = match[1]
    const quote = match[2]
    const declarationValue = match[3]
    if (!declarationName || !quote || declarationValue === undefined) {
      continue
    }
    if (new RegExp(`\\s${escapeRegExp(declarationName)}=("|')`, 'u').test(output)) {
      continue
    }
    output = output.replace(/>$/u, ` ${declarationName}=${quote}${escapeXml(declarationValue)}${quote}>`)
  }
  return output
}

function replaceRelationshipIds(xml: string, relationshipIds: ReadonlyMap<string, string>): string {
  if (relationshipIds.size === 0) {
    return xml
  }
  return xml.replace(drawingRelationshipAttributePattern, (match, quote: string, id: string) => {
    const nextId = relationshipIds.get(id)
    return nextId ? match.replace(`${quote}${id}${quote}`, `${quote}${nextId}${quote}`) : match
  })
}

function drawingMaxNonVisualPropertyId(drawingXml: string | null): number {
  if (!drawingXml) {
    return 1
  }
  return Math.max(
    1,
    ...[...drawingXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\bid="([0-9]+)"/gu)]
      .map((match) => Number(match[1]))
      .filter(Number.isSafeInteger),
  )
}

function replaceAnchorNonVisualPropertyId(anchorXml: string, nextId: number): string {
  return anchorXml.replace(nonVisualPropertyIdPattern, `$1${String(nextId)}$3`)
}

function extensionFromPath(path: string): string | null {
  const fileName = normalizeZipPath(path).slice(normalizeZipPath(path).lastIndexOf('/') + 1)
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
  const existingDefaultPattern = new RegExp(`<Default\\b[^>]*\\bExtension=("|')${escapeRegExp(extension)}\\1`, 'u')
  if (existingDefaultPattern.test(contentTypesXml) || !contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/></Types>`,
  )
}

function addWorksheetDrawing(sheetXml: string, relationshipId: string): string {
  const withNamespace = ensureRelationshipNamespace(sheetXml)
  if (/<drawing\b/u.test(withNamespace)) {
    return withNamespace.replace(/<drawing\b[^>]*(?:\/?>|>[\s\S]*?<\/drawing>)/u, `<drawing r:id="${relationshipId}"/>`)
  }
  return withNamespace.replace('</worksheet>', `<drawing r:id="${relationshipId}"/></worksheet>`)
}

function collectDrawingDependencyPaths(input: {
  readonly zip: XlsxZipEntries
  readonly drawingPath: string
  readonly drawingRelationships: readonly ParsedRelationship[]
}): Set<string> {
  const collectedPaths = new Set<string>()
  const pending = input.drawingRelationships
    .filter((relationship) => relationship.type !== chartRelationshipType && relationship.targetMode !== 'External')
    .map((relationship) => normalizeZipPath(resolveTargetPath(input.drawingPath, relationship.target)))

  collectedPaths.add(normalizeZipPath(input.drawingPath))
  const drawingRelsPath = drawingRelationshipsPath(input.drawingPath)
  if (input.zip[drawingRelsPath]) {
    collectedPaths.add(drawingRelsPath)
  }

  while (pending.length > 0) {
    const nextPartPath = pending.pop()
    if (!nextPartPath || collectedPaths.has(nextPartPath) || !input.zip[nextPartPath]) {
      continue
    }
    collectedPaths.add(nextPartPath)
    const relsPath = drawingRelationshipsPath(nextPartPath)
    const relsXml = getZipText(input.zip, relsPath)
    if (!relsXml) {
      continue
    }
    collectedPaths.add(relsPath)
    parseRelationships(relsXml).forEach((relationship) => {
      if (relationship.targetMode === 'External') {
        return
      }
      pending.push(normalizeZipPath(resolveTargetPath(nextPartPath, relationship.target)))
    })
  }

  return collectedPaths
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

function preservedPartsByPath(parts: readonly WorkbookPreservedPackagePartSnapshot[]): Map<string, Uint8Array> {
  const output = new Map<string, Uint8Array>()
  parts.forEach((part) => {
    const bytes = decodedPartBytes(part)
    if (bytes) {
      output.set(normalizeZipPath(part.path), bytes)
    }
  })
  return output
}

function preservedPartText(partsByPath: ReadonlyMap<string, Uint8Array>, path: string): string | null {
  const bytes = partsByPath.get(normalizeZipPath(path))
  return bytes ? strFromU8(bytes) : null
}

function worksheetDrawingTarget(drawingIndex: number): string {
  return `../drawings/drawing${String(drawingIndex)}.xml`
}

function nextAvailableWorksheetDrawingTarget(input: {
  readonly zip: XlsxZipEntries
  readonly usedDrawingPaths: ReadonlySet<string>
  readonly reservedDrawingPaths: ReadonlySet<string>
}): string {
  for (let drawingIndex = 1; drawingIndex < Number.MAX_SAFE_INTEGER; drawingIndex += 1) {
    const target = worksheetDrawingTarget(drawingIndex)
    const drawingPath = normalizeZipPath(`xl/drawings/drawing${String(drawingIndex)}.xml`)
    if (!input.zip[drawingPath] && !input.usedDrawingPaths.has(drawingPath) && !input.reservedDrawingPaths.has(drawingPath)) {
      return target
    }
  }
  throw new Error('Unable to allocate XLSX drawing part')
}

function worksheetDrawingPaths(zip: XlsxZipEntries, sheets: readonly WorkbookSnapshot['sheets'][number][]): Set<string> {
  const paths = new Set<string>()
  sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((_sheet, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
      const drawingRelationship = sheetRelationships.find((relationship) => relationship.type === worksheetDrawingRelationshipType)
      if (drawingRelationship) {
        paths.add(normalizeZipPath(resolveTargetPath(sheetPath, drawingRelationship.target)))
      }
    })
  return paths
}

function reservedDrawingArtifactPaths(sheets: readonly WorkbookSnapshot['sheets'][number][]): Set<string> {
  const paths = new Set<string>()
  for (const sheet of sheets) {
    const relationshipTarget = sheet.metadata?.drawingArtifacts?.relationshipTarget
    if (relationshipTarget) {
      paths.add(normalizeZipPath(resolveTargetPath('xl/worksheets/sheet1.xml', relationshipTarget)))
    }
  }
  return paths
}

function mergedDrawingResult(input: {
  readonly currentDrawingXml: string | null
  readonly preservedDrawingXml: string
  readonly currentRelationships: readonly ParsedRelationship[]
  readonly preservedRelationships: readonly ParsedRelationship[]
}): {
  readonly drawingXml: string
  readonly relationships: readonly ParsedRelationship[]
} {
  const currentAnchors = input.currentDrawingXml ? drawingAnchors(input.currentDrawingXml) : []
  const preservedNonChartRelationships = input.preservedRelationships.filter((relationship) => relationship.type !== chartRelationshipType)
  const preservedNonChartAnchors = drawingAnchors(input.preservedDrawingXml).filter((anchor) => !anchorContainsChart(anchor))

  const relationships = [...input.currentRelationships]
  const relationshipIds = new Map<string, string>()
  preservedNonChartRelationships.forEach((relationship) => {
    const nextId = nextRelationshipId(relationships)
    relationships.push({ ...relationship, id: nextId })
    relationshipIds.set(relationship.id, nextId)
  })

  let nextNonVisualPropertyId = drawingMaxNonVisualPropertyId(input.currentDrawingXml) + 1
  const rewrittenAnchors = preservedNonChartAnchors.map((anchor) => {
    const withRelationshipIds = replaceRelationshipIds(anchor, relationshipIds)
    const withObjectId = replaceAnchorNonVisualPropertyId(withRelationshipIds, nextNonVisualPropertyId)
    nextNonVisualPropertyId += 1
    return withObjectId
  })

  const baseDrawingXml = input.currentDrawingXml ?? input.preservedDrawingXml
  const currentRootOpenTag = drawingRootOpenTag(baseDrawingXml)
  const preservedRootOpenTag = drawingRootOpenTag(input.preservedDrawingXml)
  if (!currentRootOpenTag) {
    return {
      drawingXml: baseDrawingXml,
      relationships,
    }
  }
  const drawingXml = `${drawingXmlDeclaration(baseDrawingXml) ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'}${mergeNamespaceDeclarations(
    currentRootOpenTag,
    preservedRootOpenTag,
  )}${[...currentAnchors, ...rewrittenAnchors].join('')}</${drawingRootTagName(currentRootOpenTag)}>`

  return {
    drawingXml,
    relationships,
  }
}

export function readImportedWorkbookDrawingArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): {
  readonly artifacts: WorkbookDrawingArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetDrawingArtifactsSnapshot>
} {
  const zip = readXlsxZipEntries(source)
  return readImportedWorkbookDrawingArtifactsFromWorksheetRelationships(
    zip,
    sheetNames.map((sheetName, sheetIndex) => {
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      return {
        name: sheetName,
        path: sheetPath,
        drawingRelationshipId: readWorksheetDrawingRelationshipId(getZipText(zip, sheetPath)),
      }
    }),
  )
}

export function readImportedWorkbookDrawingArtifactsFromWorksheetRelationships(
  source: XlsxZipSource,
  worksheets: readonly {
    readonly name: string
    readonly path: string
    readonly drawingRelationshipId?: string | null
  }[],
): {
  readonly artifacts: WorkbookDrawingArtifactsSnapshot | undefined
  readonly sheetArtifactsByName: Map<string, WorkbookSheetDrawingArtifactsSnapshot>
} {
  const zip = readXlsxZipEntries(source)
  const sheetArtifactsByName = new Map<string, WorkbookSheetDrawingArtifactsSnapshot>()
  const partPaths = new Set<string>()

  worksheets.forEach(({ name: sheetName, path: sheetPath, drawingRelationshipId: relationshipId }) => {
    if (!relationshipId) {
      return
    }
    const sheetRelationships = parseRelationships(getZipText(zip, worksheetRelationshipsPath(sheetPath)))
    const drawingRelationship = sheetRelationships.find(
      (relationship) => relationship.id === relationshipId && relationship.type === worksheetDrawingRelationshipType,
    )
    if (!drawingRelationship) {
      return
    }
    const drawingPath = normalizeZipPath(resolveTargetPath(sheetPath, drawingRelationship.target))
    const drawingXml = getZipText(zip, drawingPath)
    if (!drawingXml) {
      return
    }
    const drawingRelationships = parseRelationships(getZipText(zip, drawingRelationshipsPath(drawingPath)))
    const hasNonChartAnchors = drawingAnchors(drawingXml).some((anchor) => !anchorContainsChart(anchor))
    const hasNonChartRelationships = drawingRelationships.some((relationship) => relationship.type !== chartRelationshipType)
    if (!hasNonChartAnchors && !hasNonChartRelationships) {
      return
    }
    sheetArtifactsByName.set(sheetName, {
      relationshipTarget: drawingRelationship.target,
    })
    collectDrawingDependencyPaths({ zip, drawingPath, drawingRelationships }).forEach((path) => {
      partPaths.add(path)
    })
  })

  if (partPaths.size === 0) {
    return {
      artifacts: undefined,
      sheetArtifactsByName,
    }
  }

  const contentTypePartPaths = [...partPaths].filter((path) => !path.endsWith('.rels'))
  const contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  const parts = [...partPaths].toSorted().flatMap((path) => {
    const bytes = zip[path]
    return bytes ? [encodedPartSnapshot(path, bytes)] : []
  })

  return {
    artifacts: {
      parts,
      ...(contentTypesXml
        ? {
            ...(readContentTypeDefaults(contentTypesXml, contentTypePartPaths).length > 0
              ? { contentTypeDefaults: readContentTypeDefaults(contentTypesXml, contentTypePartPaths) }
              : {}),
            ...(readContentTypeOverrides(contentTypesXml, contentTypePartPaths).length > 0
              ? { contentTypeOverrides: readContentTypeOverrides(contentTypesXml, contentTypePartPaths) }
              : {}),
          }
        : {}),
    },
    sheetArtifactsByName,
  }
}

function worksheetRelationshipsPath(worksheetPath: string): string {
  const normalizedPath = normalizeZipPath(worksheetPath)
  const directory = normalizedPath.slice(0, normalizedPath.lastIndexOf('/'))
  const fileName = normalizedPath.slice(normalizedPath.lastIndexOf('/') + 1)
  return `${directory}/_rels/${fileName}.rels`
}

export function addExportDrawingArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookArtifacts = snapshot.workbook.metadata?.drawingArtifacts
  const sheetsWithDrawingArtifacts = snapshot.sheets.filter((sheet) => sheet.metadata?.drawingArtifacts)
  if (!workbookArtifacts || sheetsWithDrawingArtifacts.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const partsByPath = preservedPartsByPath(workbookArtifacts.parts)
  const preservedDrawingPartPaths = new Set<string>()
  const copiedPartPaths = new Set<string>()
  const reservedDrawingPaths = reservedDrawingArtifactPaths(sheetsWithDrawingArtifacts)
  const usedDrawingPaths = worksheetDrawingPaths(zip, snapshot.sheets)
  let changed = false

  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetDrawingArtifacts = sheet.metadata?.drawingArtifacts
      if (!sheetDrawingArtifacts) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }

      const preservedDrawingPath = normalizeZipPath(resolveTargetPath(sheetPath, sheetDrawingArtifacts.relationshipTarget))
      const preservedDrawingXml = preservedPartText(partsByPath, preservedDrawingPath)
      if (!preservedDrawingXml) {
        return
      }
      const preservedDrawingRelationshipsPath = drawingRelationshipsPath(preservedDrawingPath)
      const preservedDrawingRelationships = parseRelationships(preservedPartText(partsByPath, preservedDrawingRelationshipsPath))
      preservedDrawingPartPaths.add(preservedDrawingPath)
      preservedDrawingPartPaths.add(preservedDrawingRelationshipsPath)

      const sheetRelsPath = `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
      const sheetRelationships = parseRelationships(getZipText(zip, sheetRelsPath))
      let drawingRelationship = sheetRelationships.find((relationship) => relationship.type === worksheetDrawingRelationshipType)
      let relationshipsChanged = false
      let nextSheetXml = sheetXml

      if (!drawingRelationship) {
        const drawingTarget =
          zip[preservedDrawingPath] || usedDrawingPaths.has(preservedDrawingPath)
            ? nextAvailableWorksheetDrawingTarget({ zip, usedDrawingPaths, reservedDrawingPaths })
            : sheetDrawingArtifacts.relationshipTarget
        drawingRelationship = {
          id: nextRelationshipId(sheetRelationships),
          type: worksheetDrawingRelationshipType,
          target: drawingTarget,
        }
        sheetRelationships.push(drawingRelationship)
        relationshipsChanged = true
        nextSheetXml = addWorksheetDrawing(nextSheetXml, drawingRelationship.id)
      }

      const drawingPath = normalizeZipPath(resolveTargetPath(sheetPath, drawingRelationship.target))
      usedDrawingPaths.add(drawingPath)
      const currentDrawingRelationshipsPath = drawingRelationshipsPath(drawingPath)
      const currentDrawingXml = getZipText(zip, drawingPath)
      const currentDrawingRelationships = parseRelationships(getZipText(zip, currentDrawingRelationshipsPath))
      const mergedDrawing = mergedDrawingResult({
        currentDrawingXml,
        preservedDrawingXml,
        currentRelationships: currentDrawingRelationships,
        preservedRelationships: preservedDrawingRelationships,
      })

      setZipText(zip, drawingPath, mergedDrawing.drawingXml)
      copiedPartPaths.add(drawingPath)
      changed = true
      if (mergedDrawing.relationships.length > 0 || currentDrawingXml) {
        setZipText(zip, currentDrawingRelationshipsPath, buildRelationshipsXml(mergedDrawing.relationships))
        copiedPartPaths.add(currentDrawingRelationshipsPath)
        changed = true
      }
      if (relationshipsChanged) {
        setZipText(zip, sheetRelsPath, buildRelationshipsXml(sheetRelationships))
      }
      if (nextSheetXml !== sheetXml) {
        setZipText(zip, sheetPath, nextSheetXml)
      }
    })

  partsByPath.forEach((partBytes, path) => {
    if (preservedDrawingPartPaths.has(path)) {
      return
    }
    zip[path] = partBytes
    copiedPartPaths.add(path)
    changed = true
  })

  let contentTypesXml = getZipText(zip, '[Content_Types].xml') ?? ''
  const copiedExtensions = new Set(
    [...copiedPartPaths].map(extensionFromPath).filter((extension): extension is string => Boolean(extension)),
  )
  for (const defaultEntry of workbookArtifacts.contentTypeDefaults ?? []) {
    if (!copiedExtensions.has(defaultEntry.extension)) {
      continue
    }
    contentTypesXml = addContentTypeDefault(contentTypesXml, defaultEntry.extension, defaultEntry.contentType)
  }
  for (const overrideEntry of workbookArtifacts.contentTypeOverrides ?? []) {
    if (!copiedPartPaths.has(normalizeZipPath(overrideEntry.partName))) {
      continue
    }
    contentTypesXml = addContentTypeOverride(contentTypesXml, overrideEntry.partName, overrideEntry.contentType)
  }
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const sheetDrawingArtifacts = sheet.metadata?.drawingArtifacts
      if (!sheetDrawingArtifacts) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetRelationships = parseRelationships(getZipText(zip, `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`))
      const drawingRelationship = sheetRelationships.find((relationship) => relationship.type === worksheetDrawingRelationshipType)
      if (!drawingRelationship) {
        return
      }
      const drawingPath = normalizeZipPath(resolveTargetPath(sheetPath, drawingRelationship.target))
      contentTypesXml = addContentTypeOverride(contentTypesXml, `/${drawingPath}`, drawingContentType)
    })
  if (contentTypesXml.length > 0) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
  }

  return changed ? zipSync(zip) : bytes
}
