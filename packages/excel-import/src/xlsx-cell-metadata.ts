import { strToU8, unzipSync, zipSync } from 'fflate'
import * as XLSX from 'xlsx'

import type { WorkbookCellMetadataReferenceSnapshot, WorkbookCellMetadataSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, readXlsxZipEntries, type XlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
}

interface ImportedCellMetadataReference {
  readonly address: string
  readonly cm?: string
  readonly vm?: string
}

export interface ImportedWorkbookCellMetadata {
  readonly workbookMetadata?: WorkbookCellMetadataSnapshot
  readonly refsBySheet: ReadonlyMap<string, readonly ImportedCellMetadataReference[]>
}

const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
const sheetMetadataRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata'
const sheetMetadataContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

function setZipText(zip: XlsxZipEntries, path: string, text: string): void {
  zip[normalizeZipPath(path)] = strToU8(text)
}

function readAttribute(xml: string, attributeName: string): string | null {
  const match = new RegExp(`\\s${attributeName}=(["'])([\\s\\S]*?)\\1`, 'u').exec(xml)
  return match?.[2] ?? null
}

function setXmlAttribute(tag: string, name: string, value: string): string {
  const attribute = `${name}="${escapeXml(value)}"`
  const existingAttribute = new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u')
  if (existingAttribute.test(tag)) {
    return tag.replace(existingAttribute, ` ${attribute}`)
  }
  return tag.replace(/\/?>$/u, (ending) => ` ${attribute}${ending}`)
}

function removeXmlAttribute(tag: string, name: string): string {
  return tag.replace(new RegExp(`\\s${name}=(["'])[\\s\\S]*?\\1`, 'u'), '')
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

function relationshipXml(relationship: ParsedRelationship): string {
  return `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(relationship.target)}"/>`
}

function appendRelationshipXml(xml: string | null, relationship: ParsedRelationship): string {
  if (!xml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${relationshipNamespace}">${relationshipXml(
      relationship,
    )}</Relationships>`
  }
  if (!xml.includes('</Relationships>')) {
    return xml
  }
  return xml.replace('</Relationships>', `${relationshipXml(relationship)}</Relationships>`)
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

function normalizeCellAddress(address: string): string {
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address))
  } catch {
    return address.trim().toUpperCase()
  }
}

function readWorksheetCellMetadataRefs(sheetXml: string | null): ImportedCellMetadataReference[] {
  if (!sheetXml) {
    return []
  }
  const refs: ImportedCellMetadataReference[] = []
  for (const cellTag of sheetXml.match(/<c\b[^>]*>/gu) ?? []) {
    const address = readAttribute(cellTag, 'r')
    const cm = readAttribute(cellTag, 'cm')
    const vm = readAttribute(cellTag, 'vm')
    if (!address || (!cm && !vm)) {
      continue
    }
    refs.push({
      address: normalizeCellAddress(address),
      ...(cm ? { cm } : {}),
      ...(vm ? { vm } : {}),
    })
  }
  return refs
}

export function readImportedWorkbookCellMetadataPart(zip: XlsxZipEntries): WorkbookCellMetadataSnapshot | undefined {
  const relationship = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels')).find(
    (entry) => entry.type === sheetMetadataRelationshipType,
  )
  const relationshipTarget = relationship?.target ?? 'metadata.xml'
  const relationshipPartPath = relationship ? resolveTargetPath('xl/workbook.xml', relationship.target) : 'xl/metadata.xml'
  const metadataXml = getZipText(zip, relationshipPartPath) ?? getZipText(zip, 'xl/metadata.xml')
  return metadataXml ? { relationshipTarget, metadataXml } : undefined
}

export function readImportedWorkbookCellMetadata(source: XlsxZipSource, sheetNames: readonly string[]): ImportedWorkbookCellMetadata {
  const zip = readXlsxZipEntries(source)
  const refsBySheet = new Map<string, readonly ImportedCellMetadataReference[]>()

  sheetNames.forEach((sheetName, sheetIndex) => {
    const refs = readWorksheetCellMetadataRefs(getZipText(zip, `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`))
    if (refs.length > 0) {
      refsBySheet.set(sheetName, refs)
    }
  })

  const workbookMetadata = readImportedWorkbookCellMetadataPart(zip)
  return {
    refsBySheet,
    ...(workbookMetadata ? { workbookMetadata } : {}),
  }
}

function cellMetadataReferenceSignature(cell: WorkbookSnapshot['sheets'][number]['cells'][number] | undefined): string {
  if (!cell) {
    return 'null'
  }
  return JSON.stringify({
    value: cell.value ?? null,
    formula: typeof cell.formula === 'string' ? cell.formula.trim() : null,
    format: cell.format ?? null,
  })
}

export function buildImportedCellMetadataReferenceSnapshots(
  refs: readonly ImportedCellMetadataReference[] | undefined,
  cells: readonly WorkbookSnapshot['sheets'][number]['cells'][number][],
): WorkbookCellMetadataReferenceSnapshot[] | undefined {
  if (!refs || refs.length === 0) {
    return undefined
  }
  const cellsByAddress = new Map(cells.map((cell) => [normalizeCellAddress(cell.address), cell]))
  return refs.map((ref) => ({
    ...ref,
    cellSignature: cellMetadataReferenceSignature(cellsByAddress.get(normalizeCellAddress(ref.address))),
  }))
}

function safeMetadataTarget(target: string): { readonly target: string; readonly partPath: string } {
  const partPath = normalizeZipPath(resolveTargetPath('xl/workbook.xml', target))
  return partPath === 'xl/metadata.xml' ? { target, partPath } : { target: 'metadata.xml', partPath: 'xl/metadata.xml' }
}

function replaceSheetMetadataRelationshipTarget(xml: string, target: string): string {
  return xml.replace(/<Relationship\b[^>]*\/?>/gu, (relationshipXmlText) => {
    return readAttribute(relationshipXmlText, 'Type') === sheetMetadataRelationshipType
      ? setXmlAttribute(relationshipXmlText, 'Target', target)
      : relationshipXmlText
  })
}

function ensureWorkbookMetadataRelationship(zip: XlsxZipEntries, target: string): boolean {
  const relsPath = 'xl/_rels/workbook.xml.rels'
  const relsXml = getZipText(zip, relsPath)
  const relationships = parseRelationships(relsXml)
  const existingRelationship = relationships.find((relationship) => relationship.type === sheetMetadataRelationshipType)
  if (existingRelationship?.target === target) {
    return false
  }
  if (existingRelationship) {
    setZipText(zip, relsPath, replaceSheetMetadataRelationshipTarget(relsXml ?? '', target))
    return true
  }
  setZipText(
    zip,
    relsPath,
    appendRelationshipXml(relsXml, {
      id: nextRelationshipId(relationships),
      target,
      type: sheetMetadataRelationshipType,
    }),
  )
  return true
}

function addMetadataContentTypeOverride(contentTypesXml: string | null, partPath: string): string | null {
  if (!contentTypesXml) {
    return null
  }
  const partName = `/${normalizeZipPath(partPath)}`
  const existingOverridePattern = new RegExp(`<Override\\b(?=[^>]*\\bPartName=(["'])${escapeRegExp(partName)}\\1)[^>]*/>`, 'u')
  if (existingOverridePattern.test(contentTypesXml)) {
    return contentTypesXml.replace(
      existingOverridePattern,
      `<Override PartName="${escapeXml(partName)}" ContentType="${sheetMetadataContentType}"/>`,
    )
  }
  if (!contentTypesXml.includes('</Types>')) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${escapeXml(partName)}" ContentType="${sheetMetadataContentType}"/></Types>`,
  )
}

function restoreWorkbookCellMetadata(zip: XlsxZipEntries, metadata: WorkbookCellMetadataSnapshot): boolean {
  const target = safeMetadataTarget(metadata.relationshipTarget)
  let changed = false
  const existingMetadataXml = getZipText(zip, target.partPath)
  if (existingMetadataXml !== metadata.metadataXml) {
    setZipText(zip, target.partPath, metadata.metadataXml)
    changed = true
  }
  changed = ensureWorkbookMetadataRelationship(zip, target.target) || changed

  const contentTypesXml = addMetadataContentTypeOverride(getZipText(zip, '[Content_Types].xml'), target.partPath)
  if (contentTypesXml && contentTypesXml !== getZipText(zip, '[Content_Types].xml')) {
    setZipText(zip, '[Content_Types].xml', contentTypesXml)
    changed = true
  }
  return changed
}

function restorableCellMetadataRefs(sheet: WorkbookSnapshot['sheets'][number]): WorkbookCellMetadataReferenceSnapshot[] {
  const cellsByAddress = new Map(sheet.cells.map((cell) => [normalizeCellAddress(cell.address), cell]))
  return (sheet.metadata?.cellMetadataRefs ?? []).filter(
    (ref) => cellMetadataReferenceSignature(cellsByAddress.get(normalizeCellAddress(ref.address))) === ref.cellSignature,
  )
}

function compareCellAddresses(left: string, right: string): number {
  const leftCell = XLSX.utils.decode_cell(left)
  const rightCell = XLSX.utils.decode_cell(right)
  return leftCell.r - rightCell.r || leftCell.c - rightCell.c
}

function metadataCellXml(ref: WorkbookCellMetadataReferenceSnapshot): string {
  const attributes = [
    `r="${escapeXml(normalizeCellAddress(ref.address))}"`,
    ref.cm ? `cm="${escapeXml(ref.cm)}"` : null,
    ref.vm ? `vm="${escapeXml(ref.vm)}"` : null,
  ].filter((entry): entry is string => Boolean(entry))
  return `<c ${attributes.join(' ')}/>`
}

function insertCellsIntoRowXml(rowXml: string, refs: readonly WorkbookCellMetadataReferenceSnapshot[]): string {
  const sortedRefs = refs.toSorted((left, right) => compareCellAddresses(left.address, right.address))
  if (/\/>\s*$/u.test(rowXml)) {
    return rowXml.replace(/\/>\s*$/u, `>${sortedRefs.map(metadataCellXml).join('')}</row>`)
  }

  let output = rowXml
  for (const ref of sortedRefs) {
    const refColumn = XLSX.utils.decode_cell(ref.address).c
    const insertBefore = [...output.matchAll(/<c\b[^>]*(?:\/>|>[\s\S]*?<\/c>)/gu)].find((match) => {
      const cellAddress = readAttribute(match[0], 'r')
      return cellAddress ? XLSX.utils.decode_cell(cellAddress).c > refColumn : false
    })
    if (insertBefore?.index !== undefined) {
      output = `${output.slice(0, insertBefore.index)}${metadataCellXml(ref)}${output.slice(insertBefore.index)}`
    } else {
      output = output.replace('</row>', `${metadataCellXml(ref)}</row>`)
    }
  }
  return output
}

function rowXmlForMetadataRefs(rowNumber: number, refs: readonly WorkbookCellMetadataReferenceSnapshot[]): string {
  const sortedRefs = refs.toSorted((left, right) => compareCellAddresses(left.address, right.address))
  return `<row r="${String(rowNumber)}">${sortedRefs.map(metadataCellXml).join('')}</row>`
}

function insertRowIntoSheetData(sheetXml: string, rowNumber: number, rowXml: string): string {
  const sheetDataMatch = /<sheetData\b[^>]*(?:\/>|>[\s\S]*?<\/sheetData>)/u.exec(sheetXml)
  if (!sheetDataMatch?.[0] || sheetDataMatch.index === undefined) {
    return sheetXml.replace(/<worksheet\b[^>]*>/u, (openingTag) => `${openingTag}<sheetData>${rowXml}</sheetData>`)
  }

  const sheetDataXml = sheetDataMatch[0]
  const nextSheetDataXml = /\/>\s*$/u.test(sheetDataXml)
    ? sheetDataXml.replace(/\/>\s*$/u, `>${rowXml}</sheetData>`)
    : insertRowIntoSheetDataXml(sheetDataXml, rowNumber, rowXml)
  return `${sheetXml.slice(0, sheetDataMatch.index)}${nextSheetDataXml}${sheetXml.slice(sheetDataMatch.index + sheetDataXml.length)}`
}

function insertRowIntoSheetDataXml(sheetDataXml: string, rowNumber: number, rowXml: string): string {
  const insertBefore = [...sheetDataXml.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gu)].find((match) => {
    const rowNumberText = readAttribute(match[0], 'r')
    return rowNumberText ? Number(rowNumberText) > rowNumber : false
  })
  if (insertBefore?.index !== undefined) {
    return `${sheetDataXml.slice(0, insertBefore.index)}${rowXml}${sheetDataXml.slice(insertBefore.index)}`
  }
  return sheetDataXml.replace('</sheetData>', `${rowXml}</sheetData>`)
}

function insertMissingMetadataCells(sheetXml: string, refs: readonly WorkbookCellMetadataReferenceSnapshot[]): string {
  const refsByRow = new Map<number, WorkbookCellMetadataReferenceSnapshot[]>()
  for (const ref of refs) {
    const rowNumber = XLSX.utils.decode_cell(ref.address).r + 1
    refsByRow.set(rowNumber, [...(refsByRow.get(rowNumber) ?? []), ref])
  }

  const handledRows = new Set<number>()
  let output = sheetXml.replace(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gu, (rowXml) => {
    const rowNumberText = readAttribute(rowXml, 'r')
    const rowNumber = rowNumberText ? Number(rowNumberText) : NaN
    const refsForRow = Number.isSafeInteger(rowNumber) ? refsByRow.get(rowNumber) : undefined
    if (!refsForRow) {
      return rowXml
    }
    handledRows.add(rowNumber)
    return insertCellsIntoRowXml(rowXml, refsForRow)
  })

  for (const [rowNumber, refsForRow] of [...refsByRow.entries()].toSorted((left, right) => left[0] - right[0])) {
    if (!handledRows.has(rowNumber)) {
      output = insertRowIntoSheetData(output, rowNumber, rowXmlForMetadataRefs(rowNumber, refsForRow))
    }
  }
  return output
}

function expandRangeForAddress(range: XLSX.Range | null, address: string): XLSX.Range {
  const decoded = XLSX.utils.decode_cell(address)
  if (!range) {
    return { s: { r: decoded.r, c: decoded.c }, e: { r: decoded.r, c: decoded.c } }
  }
  return {
    s: { r: Math.min(range.s.r, decoded.r), c: Math.min(range.s.c, decoded.c) },
    e: { r: Math.max(range.e.r, decoded.r), c: Math.max(range.e.c, decoded.c) },
  }
}

function updateWorksheetDimension(sheetXml: string, refs: readonly WorkbookCellMetadataReferenceSnapshot[]): string {
  let range: XLSX.Range | null = null
  const currentDimension = /<dimension\b[^>]*>/u.exec(sheetXml)?.[0]
  const currentRef = currentDimension ? readAttribute(currentDimension, 'ref') : null
  if (currentRef) {
    try {
      range = XLSX.utils.decode_range(currentRef)
    } catch {
      range = null
    }
  }
  for (const ref of refs) {
    range = expandRangeForAddress(range, ref.address)
  }
  if (!range) {
    return sheetXml
  }
  const nextDimension = currentDimension
    ? setXmlAttribute(currentDimension, 'ref', XLSX.utils.encode_range(range))
    : `<dimension ref="${escapeXml(XLSX.utils.encode_range(range))}"/>`
  return currentDimension
    ? sheetXml.replace(currentDimension, nextDimension)
    : sheetXml.replace(/<worksheet\b[^>]*>/u, (openingTag) => `${openingTag}${nextDimension}`)
}

function applyCellMetadataRefsToSheetXml(
  sheetXml: string,
  refs: readonly WorkbookCellMetadataReferenceSnapshot[],
): { readonly xml: string; readonly changed: boolean } {
  if (refs.length === 0) {
    return { xml: sheetXml, changed: false }
  }

  const refsByAddress = new Map(refs.map((ref) => [normalizeCellAddress(ref.address), ref]))
  const handledAddresses = new Set<string>()
  let output = sheetXml.replace(/<c\b[^>]*>/gu, (cellTag) => {
    const address = readAttribute(cellTag, 'r')
    const ref = address ? refsByAddress.get(normalizeCellAddress(address)) : undefined
    if (!ref) {
      return cellTag
    }
    handledAddresses.add(normalizeCellAddress(ref.address))
    const withCm = ref.cm ? setXmlAttribute(cellTag, 'cm', ref.cm) : removeXmlAttribute(cellTag, 'cm')
    return ref.vm ? setXmlAttribute(withCm, 'vm', ref.vm) : removeXmlAttribute(withCm, 'vm')
  })

  const missingRefs = refs.filter((ref) => !handledAddresses.has(normalizeCellAddress(ref.address)))
  if (missingRefs.length > 0) {
    output = insertMissingMetadataCells(output, missingRefs)
  }
  output = updateWorksheetDimension(output, refs)
  return { xml: output, changed: output !== sheetXml }
}

export function addExportCellMetadataToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const workbookCellMetadata = snapshot.workbook.metadata?.cellMetadata
  if (!workbookCellMetadata) {
    return bytes
  }

  const zip = unzipSync(bytes)
  let changed = restoreWorkbookCellMetadata(zip, workbookCellMetadata)
  snapshot.sheets
    .toSorted((left, right) => left.order - right.order)
    .forEach((sheet, sheetIndex) => {
      const refs = restorableCellMetadataRefs(sheet)
      if (refs.length === 0) {
        return
      }
      const sheetPath = `xl/worksheets/sheet${String(sheetIndex + 1)}.xml`
      const sheetXml = getZipText(zip, sheetPath)
      if (!sheetXml) {
        return
      }
      const nextSheet = applyCellMetadataRefsToSheetXml(sheetXml, refs)
      if (nextSheet.changed) {
        setZipText(zip, sheetPath, nextSheet.xml)
        changed = true
      }
    })

  return changed ? zipSync(zip) : bytes
}
