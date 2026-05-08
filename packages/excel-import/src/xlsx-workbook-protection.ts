import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'

import type { WorkbookProtectionSnapshot, WorkbookSnapshot } from '@bilig/protocol'
import { escapeXmlAttribute } from './xlsx-export-xml.js'
import { readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

type ZipEntries = Record<string, Uint8Array>

const workbookProtectionTailElements = [
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
] as const

const workbookProtectionBooleanAttributes = ['lockStructure', 'lockWindows', 'lockRevision'] as const
const xmlNamePattern = /^[A-Za-z_:][\w:.-]*$/u
const xmlNamedEntities: Readonly<Record<string, string>> = {
  amp: '&',
  apos: "'",
  gt: '>',
  lt: '<',
  quot: '"',
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

function unescapeXmlAttribute(value: string): string {
  return value.replace(
    /&#x([0-9a-fA-F]+);|&#([0-9]+);|&(quot|apos|lt|gt|amp);/gu,
    (match, hex: string | undefined, decimal: string | undefined, named: string | undefined) => {
      if (hex || decimal) {
        const codePoint = Number.parseInt(hex ?? decimal ?? '', hex ? 16 : 10)
        return Number.isSafeInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : match
      }
      return named ? (xmlNamedEntities[named] ?? match) : match
    },
  )
}

function readWorkbookProtectionXmlAttributes(workbookXml: string): WorkbookProtectionSnapshot['xmlAttributes'] | undefined {
  const match = /<workbookProtection\b[^>]*(?:\/>|>[\s\S]*?<\/workbookProtection>)/u.exec(workbookXml)
  if (!match) {
    return undefined
  }
  const attributes = [...match[0].matchAll(/\s([A-Za-z_:][\w:.-]*)=(?:"([^"]*)"|'([^']*)')/gu)].map((attributeMatch) => ({
    name: attributeMatch[1] ?? '',
    value: unescapeXmlAttribute(attributeMatch[2] ?? attributeMatch[3] ?? ''),
  }))
  return attributes.length > 0 ? attributes : undefined
}

function booleanAttributeValue(value: string | undefined): boolean | undefined {
  const normalized = value?.toLowerCase()
  if (normalized === '1' || normalized === 'true') {
    return true
  }
  if (normalized === '0' || normalized === 'false') {
    return false
  }
  return undefined
}

function xmlAttributesByName(attributes: readonly NonNullable<WorkbookProtectionSnapshot['xmlAttributes']>[number][]): Map<string, string> {
  return new Map(attributes.map((attribute) => [attribute.name, attribute.value]))
}

function setAttribute(attributes: Map<string, string>, name: string, value: string): void {
  if (!xmlNamePattern.test(name)) {
    return
  }
  attributes.set(name, value)
}

function buildWorkbookProtectionXml(protection: WorkbookProtectionSnapshot): string | null {
  const attributes = new Map<string, string>()
  for (const attribute of protection.xmlAttributes ?? []) {
    setAttribute(attributes, attribute.name, attribute.value)
  }
  for (const name of workbookProtectionBooleanAttributes) {
    const value = protection[name]
    if (typeof value === 'boolean' && !attributes.has(name)) {
      attributes.set(name, value ? '1' : '0')
    }
  }
  if (attributes.size === 0) {
    return null
  }
  return `<workbookProtection${[...attributes.entries()].map(([name, value]) => ` ${name}="${escapeXmlAttribute(value)}"`).join('')}/>`
}

function insertWorkbookProtection(workbookXml: string, workbookProtectionXml: string): string {
  if (/<workbookProtection\b/u.test(workbookXml)) {
    return workbookXml.replace(/<workbookProtection\b[^>]*(?:\/>|>[\s\S]*?<\/workbookProtection>)/u, workbookProtectionXml)
  }

  let insertIndex = workbookXml.indexOf('</workbook>')
  for (const elementName of workbookProtectionTailElements) {
    const elementIndex = workbookXml.search(new RegExp(`<${elementName}\\b`, 'u'))
    if (elementIndex >= 0 && (insertIndex < 0 || elementIndex < insertIndex)) {
      insertIndex = elementIndex
    }
  }
  if (insertIndex < 0) {
    return workbookXml
  }
  return `${workbookXml.slice(0, insertIndex)}${workbookProtectionXml}${workbookXml.slice(insertIndex)}`
}

export function addExportWorkbookProtectionToXlsxBytes(bytes: Uint8Array, snapshot: WorkbookSnapshot): Uint8Array {
  const protection = snapshot.workbook.metadata?.workbookProtection
  const workbookProtectionXml = protection ? buildWorkbookProtectionXml(protection) : null
  if (!workbookProtectionXml) {
    return bytes
  }

  const zip = unzipSync(bytes)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) {
    return bytes
  }
  setZipText(zip, 'xl/workbook.xml', insertWorkbookProtection(workbookXml, workbookProtectionXml))
  return zipSync(zip)
}

export function readImportedWorkbookProtection(source: XlsxZipSource): WorkbookProtectionSnapshot | undefined {
  const zip = readXlsxZipEntries(source)
  const workbookXml = getZipText(zip, 'xl/workbook.xml')
  if (!workbookXml || !/<workbookProtection\b/u.test(workbookXml)) {
    return undefined
  }

  const xmlAttributes = readWorkbookProtectionXmlAttributes(workbookXml)
  if (!xmlAttributes) {
    return undefined
  }
  const attributesByName = xmlAttributesByName(xmlAttributes)
  const lockStructure = booleanAttributeValue(attributesByName.get('lockStructure'))
  const lockWindows = booleanAttributeValue(attributesByName.get('lockWindows'))
  const lockRevision = booleanAttributeValue(attributesByName.get('lockRevision'))
  const protection: WorkbookProtectionSnapshot = {
    ...(lockStructure !== undefined ? { lockStructure } : {}),
    ...(lockWindows !== undefined ? { lockWindows } : {}),
    ...(lockRevision !== undefined ? { lockRevision } : {}),
    xmlAttributes,
  }
  return protection
}
