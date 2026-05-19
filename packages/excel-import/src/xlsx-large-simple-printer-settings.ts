import type { SheetMetadataSnapshot, WorkbookPrinterSettingsSnapshot } from '@bilig/protocol'
import { getZipText, normalizeZipPath, type XlsxZipEntries } from './xlsx-zip.js'
import { parseRelationships, resolveTargetPath } from './xlsx-pivot-artifacts.js'

export type PrintPageSetupSnapshot = NonNullable<SheetMetadataSnapshot['printPageSetup']>

const binaryChunkSize = 0x8000
const printerSettingsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
const printerSettingsPathPattern = /^xl\/printerSettings\/[^/]+\.bin$/u
const printPageSetupElementNames = ['printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks'] as const

type PrintPageSetupElementName = (typeof printPageSetupElementNames)[number]

export function readLargeSimpleSheetPrintMetadata(
  zip: XlsxZipEntries,
  worksheetPath: string,
  printPageSetup: PrintPageSetupSnapshot | undefined,
): Pick<SheetMetadataSnapshot, 'printerSettings' | 'printPageSetup'> | null {
  const printerSettings = readLargeSimpleSheetPrinterSettings(zip, worksheetPath, printPageSetup)
  if (printerSettings === null) {
    return null
  }
  return {
    ...(printerSettings ? { printerSettings } : {}),
    ...(printPageSetup ? { printPageSetup } : {}),
  }
}

export function readLargeSimpleSheetPrintPageSetup(worksheetXml: string): PrintPageSetupSnapshot | undefined {
  const printPageSetup: PrintPageSetupSnapshot = {}
  for (const elementName of printPageSetupElementNames) {
    const xml = readElementXml(worksheetXml, elementName)
    if (xml) {
      appendLargeSimplePrintPageSetupElement(printPageSetup, elementName, xml)
    }
  }
  return Object.keys(printPageSetup).length > 0 ? printPageSetup : undefined
}

export function isLargeSimplePrintPageSetupElementName(value: string): value is PrintPageSetupElementName {
  switch (value) {
    case 'printOptions':
    case 'pageMargins':
    case 'pageSetup':
    case 'headerFooter':
    case 'rowBreaks':
    case 'colBreaks':
      return true
    default:
      return false
  }
}

export function appendLargeSimplePrintPageSetupElement(
  printPageSetup: PrintPageSetupSnapshot,
  elementName: PrintPageSetupElementName,
  xml: string,
): void {
  switch (elementName) {
    case 'printOptions':
      printPageSetup.printOptionsXml = xml
      break
    case 'pageMargins':
      printPageSetup.pageMarginsXml = xml
      break
    case 'pageSetup':
      printPageSetup.pageSetupXml = xml
      break
    case 'headerFooter':
      printPageSetup.headerFooterXml = xml
      break
    case 'rowBreaks':
      printPageSetup.rowBreaksXml = xml
      break
    case 'colBreaks':
      printPageSetup.colBreaksXml = xml
      break
  }
}

function readLargeSimpleSheetPrinterSettings(
  zip: XlsxZipEntries,
  worksheetPath: string,
  printPageSetup: PrintPageSetupSnapshot | undefined,
): WorkbookPrinterSettingsSnapshot[] | null | undefined {
  const relationships = parseRelationships(getZipText(zip, worksheetRelationshipsPath(worksheetPath))).filter(
    (relationship) => relationship.type === printerSettingsRelationshipType || relationship.type.endsWith('/printerSettings'),
  )
  if (relationships.length === 0) {
    return undefined
  }

  const settings: WorkbookPrinterSettingsSnapshot[] = []
  for (const relationship of relationships) {
    const partPath = normalizeZipPath(resolveTargetPath(worksheetPath, relationship.target))
    if (!printerSettingsPathPattern.test(partPath)) {
      return null
    }
    const bytes = zip[partPath]
    if (!bytes) {
      return null
    }
    const pageSetupXml = readPageSetupXml(printPageSetup?.pageSetupXml, relationship.id)
    settings.push({
      relationshipTarget: relationship.target,
      storage: 'base64',
      dataBase64: encodeBase64(bytes),
      byteLength: bytes.byteLength,
      ...(pageSetupXml ? { pageSetupXml } : {}),
    })
  }
  return settings.length > 0 ? settings : undefined
}

function readPageSetupXml(pageSetupXml: string | undefined, relationshipId: string): string | undefined {
  if (!pageSetupXml) {
    return undefined
  }
  const pageSetupRelationshipId = readXmlAttribute(pageSetupXml, 'r:id') ?? readXmlAttribute(pageSetupXml, 'id')
  return pageSetupRelationshipId === relationshipId ? pageSetupXml : undefined
}

function readElementXml(worksheetXml: string, elementName: PrintPageSetupElementName): string | undefined {
  return new RegExp(`<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b[^>]*(?:/>|>[\\s\\S]*?</\\1>)`, 'u').exec(worksheetXml)?.[0]
}

function worksheetRelationshipsPath(worksheetPath: string): string {
  const directory = worksheetPath.slice(0, worksheetPath.lastIndexOf('/'))
  const fileName = worksheetPath.slice(worksheetPath.lastIndexOf('/') + 1)
  return `${directory}/_rels/${fileName}.rels`
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function encodeBinaryString(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += binaryChunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + binaryChunkSize))
  }
  return binary
}

function encodeBase64(bytes: Uint8Array): string {
  const btoa = globalThis.btoa
  if (typeof btoa === 'function') {
    return btoa(encodeBinaryString(bytes))
  }
  return Buffer.from(bytes).toString('base64')
}
