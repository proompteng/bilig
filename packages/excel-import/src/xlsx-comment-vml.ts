import { strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookCommentThreadSnapshot, WorkbookLegacyCommentVmlSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

export interface ImportedLegacyCommentVml {
  readonly relationshipTarget: string
  readonly vmlXml: string
}

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'

const legacyDrawingTailElements = [
  'drawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
] as const

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const id = readAttribute(attributes, 'Id')
    const target = readAttribute(attributes, 'Target')
    const type = readAttribute(attributes, 'Type')
    return id && target && type ? [{ id, target, type }] : []
  })
}

function buildRelationshipsXml(relationships: readonly ParsedRelationship[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    `<Relationships xmlns="${relationshipNamespace}">`,
    ...relationships.map(
      (relationship) =>
        `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(
          relationship.target,
        )}"/>`,
    ),
    '</Relationships>',
  ].join('')
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

function nextPartIndex(zip: XlsxZipEntries, prefix: string, suffix: string): number {
  let next = 1
  for (const path of Object.keys(zip)) {
    if (!path.startsWith(prefix) || !path.endsWith(suffix)) {
      continue
    }
    const raw = path.slice(prefix.length, -suffix.length)
    const value = Number(raw)
    if (Number.isInteger(value) && value >= next) {
      next = value + 1
    }
  }
  return next
}

function resolveTargetPath(basePartPath: string, target: string): string {
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

function readLegacyDrawingRelationshipId(sheetXml: string | null): string | null {
  if (!sheetXml) {
    return null
  }
  const tag = /<legacyDrawing\b[^>]*>/u.exec(sheetXml)?.[0]
  return tag ? (readAttribute(tag, 'r:id') ?? readAttribute(tag, 'id')) : null
}

function isLegacyCommentVmlXml(vmlXml: string): boolean {
  return /<x:ClientData\b[^>]*\bObjectType=(["'])Note\1/u.test(vmlXml)
}

function readSheetRelationshipPath(sheetIndex: number): string {
  return `xl/worksheets/_rels/sheet${String(sheetIndex + 1)}.xml.rels`
}

function normalizeCommentAddress(address: string): string {
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address))
  } catch {
    return address.trim().toUpperCase()
  }
}

export function legacyCommentThreadSignature(commentThreads: readonly WorkbookCommentThreadSnapshot[] | undefined): string {
  const normalized = (commentThreads ?? [])
    .map((thread) => ({
      sheetName: thread.sheetName,
      address: normalizeCommentAddress(thread.address),
      comments: thread.comments.map((comment) => ({
        body: comment.body,
        authorDisplayName: comment.authorDisplayName ?? '',
      })),
    }))
    .toSorted((left, right) => `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`))
  return JSON.stringify(normalized)
}

export function readImportedWorkbookLegacyCommentVml(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, ImportedLegacyCommentVml> {
  const zip = readXlsxZipEntries(source)
  const legacyCommentVmlBySheet = new Map<string, ImportedLegacyCommentVml>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const relationshipId = readLegacyDrawingRelationshipId(getZipText(zip, sheetPath))
    if (!relationshipId) {
      return
    }
    const relationship = parseRelationships(getZipText(zip, readSheetRelationshipPath(sheetIndex))).find(
      (entry) => entry.id === relationshipId && entry.type === vmlDrawingRelationshipType,
    )
    if (!relationship) {
      return
    }
    const vmlXml = getZipText(zip, resolveTargetPath(sheetPath, relationship.target))
    if (!vmlXml || !isLegacyCommentVmlXml(vmlXml)) {
      return
    }
    legacyCommentVmlBySheet.set(sheetName, {
      relationshipTarget: relationship.target,
      vmlXml,
    })
  })

  return legacyCommentVmlBySheet
}

function ensureOfficeRelationshipNamespace(sheetXml: string): string {
  if (/\sxmlns:r=(["'])[\s\S]*?\1/u.test(sheetXml)) {
    return sheetXml
  }
  return sheetXml.replace(/<worksheet\b([^>]*)>/u, `<worksheet$1 xmlns:r="${officeRelationshipNamespace}">`)
}

function insertLegacyDrawing(sheetXml: string, relationshipId: string): string {
  const withNamespace = ensureOfficeRelationshipNamespace(sheetXml)
  const legacyDrawingXml = `<legacyDrawing r:id="${escapeXml(relationshipId)}"/>`
  if (/<legacyDrawing\b[^>]*(?:\/>|>[\s\S]*?<\/legacyDrawing>)/u.test(withNamespace)) {
    return withNamespace.replace(/<legacyDrawing\b[^>]*(?:\/>|>[\s\S]*?<\/legacyDrawing>)/u, legacyDrawingXml)
  }

  let insertIndex = withNamespace.indexOf('</worksheet>')
  for (const elementName of legacyDrawingTailElements) {
    const elementIndex = withNamespace.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return withNamespace
  }
  return `${withNamespace.slice(0, insertIndex)}${legacyDrawingXml}${withNamespace.slice(insertIndex)}`
}

function addVmlContentTypeDefault(contentTypesXml: string | null): string | null {
  if (!contentTypesXml) {
    return null
  }
  if (/<Default\b[^>]*\bExtension=(["'])vml\1/u.test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace('</Types>', `<Default Extension="vml" ContentType="${vmlDrawingContentType}"/></Types>`)
}

function preserveLegacyCommentVmlForSheet(input: {
  zip: XlsxZipEntries
  sheet: WorkbookSnapshot['sheets'][number]
  sheetIndex: number
  preserved: WorkbookLegacyCommentVmlSnapshot
}): boolean {
  if (legacyCommentThreadSignature(input.sheet.metadata?.commentThreads) !== input.preserved.commentSignature) {
    return false
  }

  const sheetPath = `xl/worksheets/sheet${String(input.sheetIndex + 1)}.xml`
  const sheetXml = getZipText(input.zip, sheetPath)
  if (!sheetXml) {
    return false
  }

  const sheetRelsPath = readSheetRelationshipPath(input.sheetIndex)
  const relationships = parseRelationships(getZipText(input.zip, sheetRelsPath))
  const existingRelationshipId = readLegacyDrawingRelationshipId(sheetXml)
  const existingRelationship = existingRelationshipId
    ? relationships.find((entry) => entry.id === existingRelationshipId && entry.type === vmlDrawingRelationshipType)
    : undefined

  if (existingRelationship) {
    setZipText(input.zip, resolveTargetPath(sheetPath, existingRelationship.target), input.preserved.vmlXml)
    return true
  }

  const relationshipId = nextRelationshipId(relationships)
  const vmlDrawingIndex = nextPartIndex(input.zip, 'xl/drawings/vmlDrawing', '.vml')
  const vmlPath = `xl/drawings/vmlDrawing${String(vmlDrawingIndex)}.vml`
  relationships.push({
    id: relationshipId,
    type: vmlDrawingRelationshipType,
    target: `../drawings/vmlDrawing${String(vmlDrawingIndex)}.vml`,
  })
  setZipText(input.zip, vmlPath, input.preserved.vmlXml)
  setZipText(input.zip, sheetRelsPath, buildRelationshipsXml(relationships))
  setZipText(input.zip, sheetPath, insertLegacyDrawing(sheetXml, relationshipId))

  const contentTypesXml = addVmlContentTypeDefault(getZipText(input.zip, '[Content_Types].xml'))
  if (contentTypesXml) {
    setZipText(input.zip, '[Content_Types].xml', contentTypesXml)
  }
  return true
}

export function addExportLegacyCommentVmlToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheetsWithPreservedVml = snapshot.sheets.filter((sheet) => sheet.metadata?.legacyCommentVml)
  if (sheetsWithPreservedVml.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = false
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const preserved = sheet.metadata?.legacyCommentVml
      if (!preserved) {
        return
      }
      changed = preserveLegacyCommentVmlForSheet({ zip, sheet, sheetIndex, preserved }) || changed
    })

  return changed ? zipSync(zip) : bytes
}
