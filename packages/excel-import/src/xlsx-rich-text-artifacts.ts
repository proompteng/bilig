import { unzipSync, zipSync } from 'fflate'

import type { WorkbookRichTextCellSnapshot, WorkbookSheetRichTextArtifactsSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'
import {
  addContentTypeOverride,
  buildRelationshipsXml,
  nextRelationshipId,
  parseRelationships,
  resolveTargetPath,
  setZipText,
} from './xlsx-pivot-artifacts.js'
import { setXmlAttribute } from './xlsx-export-xml.js'

const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const sharedStringsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const sharedStringsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

interface SharedStringEntry {
  readonly text: string
  readonly xml: string
  readonly rich: boolean
}

interface ExportRichTextCell {
  readonly address: string
  readonly storage: WorkbookRichTextCellSnapshot['storage']
  readonly xml: string
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
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

function decodeXmlAttribute(value: string): string {
  return decodeXmlText(value)
}

function readWorkbookSheetEntries(workbookXml: string | null): WorkbookSheetEntry[] {
  if (!workbookXml) {
    return []
  }
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b([^>]*)\/?>/gu)].flatMap((match) => {
    const attributes = match[1] ?? ''
    const rawName = readXmlAttribute(attributes, 'name')
    const relationshipId = readXmlAttribute(attributes, 'r:id') ?? readXmlAttribute(attributes, 'id')
    return rawName && relationshipId ? [{ name: decodeXmlAttribute(rawName), relationshipId }] : []
  })
}

function worksheetPathsBySheetName(zip: XlsxZipEntries, sheetNames: readonly string[]): Map<string, string> {
  const paths = new Map<string, string>()
  const relationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const worksheetRelationshipsById = new Map(
    relationships
      .filter((relationship) => relationship.type === worksheetRelationshipType || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, normalizeZipPath(resolveTargetPath('xl/workbook.xml', relationship.target))]),
  )
  readWorkbookSheetEntries(getZipText(zip, 'xl/workbook.xml')).forEach((entry) => {
    const path = worksheetRelationshipsById.get(entry.relationshipId)
    if (path) {
      paths.set(entry.name, path)
    }
  })
  sheetNames.forEach((sheetName, index) => {
    if (!paths.has(sheetName)) {
      paths.set(sheetName, `xl/worksheets/sheet${String(index + 1)}.xml`)
    }
  })
  return paths
}

function hasRichTextRuns(xml: string): boolean {
  return /<(?:[A-Za-z_][\w.-]*:)?r\b[^>]*>/u.test(xml)
}

function stringItemText(xml: string): string {
  return [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
    .map((match) => decodeXmlText(match[1] ?? ''))
    .join('')
}

function readElementText(xml: string, elementName: 'v'): string | null {
  return (
    new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>([\\s\\S]*?)</(?:[A-Za-z_][\\w.-]*:)?${elementName}>`, 'u').exec(xml)?.[1] ??
    null
  )
}

function readStringElement(xml: string, elementName: 'is' | 'si'): string | null {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(xml)?.[0] ?? null
}

function readSharedStringEntries(zip: XlsxZipEntries): SharedStringEntry[] {
  const sharedStringsXml = getZipText(zip, 'xl/sharedStrings.xml')
  if (!sharedStringsXml) {
    return []
  }
  return [...sharedStringsXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?si)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)/gu)].map((match) => {
    const xml = match[0]
    return {
      text: stringItemText(xml),
      xml,
      rich: hasRichTextRuns(xml),
    }
  })
}

function readSharedStringIndex(cellXml: string): number | null {
  const rawValue = readElementText(cellXml, 'v')?.trim()
  if (!rawValue) {
    return null
  }
  const index = Number(decodeXmlText(rawValue))
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function richTextArtifactsForWorksheet(
  sheetXml: string | null,
  sharedStrings: readonly SharedStringEntry[],
): WorkbookSheetRichTextArtifactsSnapshot | undefined {
  if (!sheetXml) {
    return undefined
  }
  const cells: WorkbookRichTextCellSnapshot[] = []
  for (const match of sheetXml.matchAll(/<((?:[A-Za-z_][\w.-]*:)?c)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)/gu)) {
    const cellXml = match[0]
    const openingTag = /<((?:[A-Za-z_][\w.-]*:)?c)\b[^>]*(?:\/>|>)/u.exec(cellXml)?.[0]
    const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
    if (!address) {
      continue
    }
    const cellType = openingTag ? readXmlAttribute(openingTag, 't') : null
    if (cellType === 's') {
      const sharedString = sharedStrings[readSharedStringIndex(cellXml) ?? -1]
      if (sharedString?.rich) {
        cells.push({
          address,
          text: sharedString.text,
          storage: 'sharedString',
          xml: sharedString.xml,
        })
      }
      continue
    }
    if (cellType === 'inlineStr') {
      const inlineStringXml = readStringElement(cellXml, 'is')
      if (inlineStringXml && hasRichTextRuns(inlineStringXml)) {
        cells.push({
          address,
          text: stringItemText(inlineStringXml),
          storage: 'inlineString',
          xml: inlineStringXml,
        })
      }
    }
  }
  return cells.length > 0 ? { cells } : undefined
}

export function readImportedWorkbookRichTextArtifacts(
  source: XlsxZipSource,
  sheetNames: readonly string[],
): Map<string, WorkbookSheetRichTextArtifactsSnapshot> {
  const zip = readXlsxZipEntries(source)
  const sharedStrings = readSharedStringEntries(zip)
  const worksheetPaths = worksheetPathsBySheetName(zip, sheetNames)
  const artifactsBySheet = new Map<string, WorkbookSheetRichTextArtifactsSnapshot>()
  sheetNames.forEach((sheetName) => {
    const worksheetPath = worksheetPaths.get(sheetName)
    const artifacts = worksheetPath ? richTextArtifactsForWorksheet(getZipText(zip, worksheetPath), sharedStrings) : undefined
    if (artifacts) {
      artifactsBySheet.set(sheetName, artifacts)
    }
  })
  return artifactsBySheet
}

function snapshotStringCellsByAddress(
  sheet: WorkbookSnapshot['sheets'][number],
): Map<string, WorkbookSnapshot['sheets'][number]['cells'][number]> {
  const cells = new Map<string, WorkbookSnapshot['sheets'][number]['cells'][number]>()
  sheet.cells.forEach((cell) => {
    if (typeof cell.value === 'string' && !cell.formula) {
      cells.set(cell.address, cell)
    }
  })
  return cells
}

function exportRichTextCells(sheet: WorkbookSnapshot['sheets'][number]): ExportRichTextCell[] {
  const textCells = snapshotStringCellsByAddress(sheet)
  return (sheet.metadata?.richTextArtifacts?.cells ?? []).flatMap((artifact): ExportRichTextCell[] => {
    const cell = textCells.get(artifact.address)
    return cell?.value === artifact.text
      ? [
          {
            address: artifact.address,
            storage: artifact.storage,
            xml: artifact.xml,
          },
        ]
      : []
  })
}

function sharedStringItemCount(sharedStringsXml: string | null): number {
  return sharedStringsXml ? [...sharedStringsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?si\b/gu)].length : 0
}

function updateSharedStringCounts(sharedStringsXml: string): string {
  const count = sharedStringItemCount(sharedStringsXml)
  return sharedStringsXml.replace(/<sst\b[^>]*>/u, (openingTag) =>
    setXmlAttribute(setXmlAttribute(openingTag, 'count', String(count)), 'uniqueCount', String(count)),
  )
}

function appendSharedStringItems(sharedStringsXml: string | null, items: readonly string[]): string {
  const baseXml =
    sharedStringsXml ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>'
  const withItems = baseXml.includes('</sst>') ? baseXml.replace('</sst>', () => `${items.join('')}</sst>`) : baseXml
  return updateSharedStringCounts(withItems)
}

function ensureSharedStringsPackageLinks(zip: XlsxZipEntries): void {
  const relationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  if (!relationships.some((relationship) => relationship.type === sharedStringsRelationshipType)) {
    relationships.push({
      id: nextRelationshipId(relationships),
      type: sharedStringsRelationshipType,
      target: 'sharedStrings.xml',
    })
    setZipText(zip, 'xl/_rels/workbook.xml.rels', buildRelationshipsXml(relationships))
  }
  const contentTypesXml = getZipText(zip, '[Content_Types].xml')
  if (contentTypesXml) {
    setZipText(zip, '[Content_Types].xml', addContentTypeOverride(contentTypesXml, '/xl/sharedStrings.xml', sharedStringsContentType))
  }
}

function cellOpeningTag(cellXml: string): string | null {
  return /<((?:[A-Za-z_][\w.-]*:)?c)\b[^>]*(?:\/>|>)/u.exec(cellXml)?.[0] ?? null
}

function cellTagName(openingTag: string): string {
  return /^<([^\s>/]+)/u.exec(openingTag)?.[1] ?? 'c'
}

function cellWithBody(cellXml: string, type: string, bodyXml: string): string {
  const openingTag = cellOpeningTag(cellXml)
  if (!openingTag) {
    return cellXml
  }
  const tagName = cellTagName(openingTag)
  const expandedOpeningTag = openingTag.replace(/\s+t=(["'])[\s\S]*?\1/u, '').replace(/\/>$/u, '>')
  return `${setXmlAttribute(expandedOpeningTag, 't', type)}${bodyXml}</${tagName}>`
}

function applyRichTextCellsToWorksheetXml(
  sheetXml: string,
  cells: readonly ExportRichTextCell[],
  sharedStringIndexes: ReadonlyMap<string, number>,
): { readonly xml: string; readonly changed: boolean } {
  if (cells.length === 0) {
    return { xml: sheetXml, changed: false }
  }
  const cellsByAddress = new Map(cells.map((cell) => [cell.address, cell]))
  let changed = false
  const xml = sheetXml.replace(/<((?:[A-Za-z_][\w.-]*:)?c)\b[^>]*(?:\/>|>[\s\S]*?<\/\1>)/gu, (cellXml) => {
    const openingTag = cellOpeningTag(cellXml)
    const address = openingTag ? readXmlAttribute(openingTag, 'r') : null
    const richTextCell = address ? cellsByAddress.get(address) : undefined
    if (!address || !richTextCell) {
      return cellXml
    }
    changed = true
    if (richTextCell.storage === 'sharedString') {
      const sharedStringIndex = sharedStringIndexes.get(address)
      return sharedStringIndex === undefined ? cellXml : cellWithBody(cellXml, 's', `<v>${String(sharedStringIndex)}</v>`)
    }
    return cellWithBody(cellXml, 'inlineStr', richTextCell.xml)
  })
  return { xml, changed }
}

export function addExportRichTextArtifactsToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const sheetsWithRichText = snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .map((sheet, sheetIndex) => ({
      sheet,
      sheetIndex,
      cells: exportRichTextCells(sheet),
    }))
    .filter((entry) => entry.cells.length > 0)
  if (sheetsWithRichText.length === 0) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const sharedStringIndexesBySheetAddress = new Map<string, number>()
  const sharedStringItems: string[] = []
  let nextSharedStringIndex = sharedStringItemCount(getZipText(zip, 'xl/sharedStrings.xml'))
  sheetsWithRichText.forEach(({ sheetIndex, cells }) => {
    cells.forEach((cell) => {
      if (cell.storage !== 'sharedString') {
        return
      }
      sharedStringIndexesBySheetAddress.set(`${String(sheetIndex)}\u0000${cell.address}`, nextSharedStringIndex)
      nextSharedStringIndex += 1
      sharedStringItems.push(cell.xml)
    })
  })
  if (sharedStringItems.length > 0) {
    setZipText(zip, 'xl/sharedStrings.xml', appendSharedStringItems(getZipText(zip, 'xl/sharedStrings.xml'), sharedStringItems))
    ensureSharedStringsPackageLinks(zip)
  }

  let changed = sharedStringItems.length > 0
  sheetsWithRichText.forEach(({ sheetIndex, cells }) => {
    const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
    const sheetXml = getZipText(zip, sheetPath)
    if (!sheetXml) {
      return
    }
    const sharedStringIndexes = new Map(
      [...sharedStringIndexesBySheetAddress.entries()]
        .filter(([key]) => key.startsWith(`${String(sheetIndex)}\u0000`))
        .map(([key, index]) => [key.slice(String(sheetIndex).length + 1), index]),
    )
    const result = applyRichTextCellsToWorksheetXml(sheetXml, cells, sharedStringIndexes)
    if (result.changed) {
      setZipText(zip, sheetPath, result.xml)
      changed = true
    }
  })

  return changed ? zipSync(zip) : bytes
}
