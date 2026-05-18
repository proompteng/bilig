import * as XLSX from 'xlsx'

import { toLiteralInput } from './workbook-import-helpers.js'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'
import { getZipText, readXlsxZipEntries, type XlsxZipSource } from './xlsx-zip.js'

interface SharedStringEntry {
  readonly text: string
}

const cellElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?c)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const cellOpeningTagPattern = /<(?:[A-Za-z_][\w.-]*:)?c\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u
const sharedStringElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?si\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?si)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const cellAddressPattern = /^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$/u
const underscoreCode = '_'.charCodeAt(0)
const lowerXCode = 'x'.charCodeAt(0)
const upperXCode = 'X'.charCodeAt(0)
const escapedNumericEntityMarker = new TextEncoder().encode('&amp;#')

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
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

function normalizeWorksheetText(value: string): string {
  const literal = toLiteralInput(decodeExcelEscapedText(value))
  return typeof literal === 'string' ? literal : value
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function normalizeAddress(address: string): string | null {
  if (!cellAddressPattern.test(address)) {
    return null
  }
  try {
    return XLSX.utils.encode_cell(XLSX.utils.decode_cell(address.replaceAll('$', '')))
  } catch {
    return null
  }
}

function isHexByte(value: number | undefined): boolean {
  return value !== undefined && ((value >= 48 && value <= 57) || (value >= 65 && value <= 70) || (value >= 97 && value <= 102))
}

function hasExcelEscapedText(bytes: Uint8Array | undefined): boolean {
  if (!bytes) {
    return false
  }
  for (let index = 0; index <= bytes.byteLength - 7; index += 1) {
    if (
      bytes[index] === underscoreCode &&
      (bytes[index + 1] === lowerXCode || bytes[index + 1] === upperXCode) &&
      isHexByte(bytes[index + 2]) &&
      isHexByte(bytes[index + 3]) &&
      isHexByte(bytes[index + 4]) &&
      isHexByte(bytes[index + 5]) &&
      bytes[index + 6] === underscoreCode
    ) {
      return true
    }
  }
  return false
}

function hasByteSequence(bytes: Uint8Array | undefined, sequence: Uint8Array): boolean {
  if (!bytes || bytes.byteLength < sequence.byteLength) {
    return false
  }
  for (let index = 0; index <= bytes.byteLength - sequence.byteLength; index += 1) {
    let matched = true
    for (let offset = 0; offset < sequence.byteLength; offset += 1) {
      if (bytes[index + offset] !== sequence[offset]) {
        matched = false
        break
      }
    }
    if (matched) {
      return true
    }
  }
  return false
}

function hasTextOverrideMarker(bytes: Uint8Array | undefined): boolean {
  return hasExcelEscapedText(bytes) || hasByteSequence(bytes, escapedNumericEntityMarker)
}

function stringItemText(xml: string): string {
  const text = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
    .map((match) => decodeXmlText(match[1] ?? ''))
    .join('')
  return normalizeWorksheetText(text)
}

function readElementText(xml: string, elementName: 'v'): string | null {
  return (
    new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b[^>]*>([\\s\\S]*?)</(?:[A-Za-z_][\\w.-]*:)?${elementName}>`, 'u').exec(xml)?.[1] ??
    null
  )
}

function readStringElement(xml: string, elementName: 'is'): string | null {
  const elementPattern = new RegExp(
    `<(?:[A-Za-z_][\\w.-]*:)?${elementName}\\b(?:[^>"']|"[^"]*"|'[^']*')*/>|<((?:[A-Za-z_][\\w.-]*:)?${elementName})\\b(?:[^>"']|"[^"]*"|'[^']*')*>[\\s\\S]*?</\\1>`,
    'u',
  )
  return elementPattern.exec(xml)?.[0] ?? null
}

function readSharedStringEntries(source: XlsxZipSource): SharedStringEntry[] {
  const sharedStringsXml = getZipText(readXlsxZipEntries(source), 'xl/sharedStrings.xml')
  if (!sharedStringsXml) {
    return []
  }
  return [...sharedStringsXml.matchAll(sharedStringElementPattern)].map((match) => ({
    text: stringItemText(match[0]),
  }))
}

function readSharedStringIndex(cellXml: string): number | null {
  const rawValue = readElementText(cellXml, 'v')?.trim()
  if (!rawValue) {
    return null
  }
  const index = Number(decodeXmlText(rawValue))
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function readCellTextValue(cellXml: string, sharedStrings: readonly SharedStringEntry[]): string | null {
  const openingTag = cellOpeningTagPattern.exec(cellXml)?.[0]
  const cellType = openingTag ? readXmlAttribute(openingTag, 't') : null
  if (cellType === 's') {
    return sharedStrings[readSharedStringIndex(cellXml) ?? -1]?.text ?? null
  }
  if (cellType === 'inlineStr') {
    const inlineStringXml = readStringElement(cellXml, 'is')
    return inlineStringXml ? stringItemText(inlineStringXml) : null
  }
  if (cellType === 'str') {
    const value = readElementText(cellXml, 'v')
    return value === null ? null : normalizeWorksheetText(decodeXmlText(value))
  }
  return null
}

function readWorksheetTextValues(sheetXml: string | null, sharedStrings: readonly SharedStringEntry[]): Map<string, string> {
  const values = new Map<string, string>()
  if (!sheetXml) {
    return values
  }
  for (const match of sheetXml.matchAll(cellElementPattern)) {
    const cellXml = match[0]
    const openingTag = cellOpeningTagPattern.exec(cellXml)?.[0]
    const rawAddress = openingTag ? readXmlAttribute(openingTag, 'r') : null
    const address = rawAddress ? normalizeAddress(rawAddress) : null
    if (!address) {
      continue
    }
    const text = readCellTextValue(cellXml, sharedStrings)
    if (text !== null) {
      values.set(address, text)
    }
  }
  return values
}

export function readImportedWorksheetTextValues(
  source: XlsxZipSource,
  sheetNames: readonly string[],
  sheetPathsByName: ReadonlyMap<string, string>,
  fallbackSheetPaths: readonly string[],
): Map<string, Map<string, string>> {
  const zip = readXlsxZipEntries(source)
  const sharedStrings = readSharedStringEntries(zip)
  const valuesBySheet = new Map<string, Map<string, string>>()
  sheetNames.forEach((sheetName, index) => {
    const sheetPath = sheetPathsByName.get(sheetName) ?? fallbackSheetPaths[index] ?? `xl/worksheets/sheet${String(index + 1)}.xml`
    const values = readWorksheetTextValues(getZipText(zip, sheetPath), sharedStrings)
    if (values.size > 0) {
      valuesBySheet.set(sheetName, values)
    }
  })
  return valuesBySheet
}

export function readImportedWorksheetTextValuesForSheet(source: XlsxZipSource, sheetPath: string): Map<string, string> {
  const zip = readXlsxZipEntries(source)
  return readWorksheetTextValues(getZipText(zip, sheetPath), readSharedStringEntries(zip))
}

export function shouldReadImportedWorksheetTextValuesForSheet(source: XlsxZipSource, sheetPath: string): boolean {
  const zip = readXlsxZipEntries(source)
  const sharedStrings = zip['xl/sharedStrings.xml']
  if (hasTextOverrideMarker(sharedStrings)) {
    return true
  }
  return !sharedStrings && hasTextOverrideMarker(zip[sheetPath])
}
