import type { WorkbookHyperlinkSnapshot } from '@bilig/protocol'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'
import { parseRelationships } from './xlsx-pivot-artifacts.js'

const hyperlinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const hyperlinkElementPattern = /<(?:[A-Za-z_][\w.-]*:)?hyperlink\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const maxExpandedHyperlinkRangeCells = 1_024
const lessThan = 60
const slash = 47
const colon = 58
const doubleQuote = 34
const singleQuote = 39
const greaterThan = 62

export interface LargeSimpleHyperlinkRef {
  readonly ref: string
  readonly relationshipId?: string
  readonly location?: string
  readonly tooltip?: string
  readonly display?: string
}

export function readLargeSimpleSheetHyperlinks(
  zip: XlsxZipEntries,
  sheetName: string,
  worksheetPath: string,
  worksheetXml: string,
): WorkbookHyperlinkSnapshot[] | null | undefined {
  const refs = readLargeSimpleSheetHyperlinkRefs(worksheetXml)
  return refs === null ? null : resolveLargeSimpleSheetHyperlinks(zip, sheetName, worksheetPath, refs)
}

export function resolveLargeSimpleSheetHyperlinks(
  zip: XlsxZipEntries,
  sheetName: string,
  worksheetPath: string,
  refs: readonly LargeSimpleHyperlinkRef[],
): WorkbookHyperlinkSnapshot[] | null | undefined {
  if (refs.length === 0) {
    return undefined
  }
  const relationships = new Map(
    parseRelationships(getZipText(zip, worksheetRelationshipsPath(worksheetPath)))
      .filter((relationship) => relationship.type === hyperlinkRelationshipType || relationship.type.endsWith('/hyperlink'))
      .map((relationship) => [relationship.id, relationship]),
  )
  const hyperlinks: WorkbookHyperlinkSnapshot[] = []
  for (const ref of refs) {
    const addresses = hyperlinkAddresses(ref.ref)
    if (!addresses) {
      return null
    }
    const relationshipTarget = ref.relationshipId ? relationships.get(ref.relationshipId)?.target : undefined
    const target = relationshipTarget ? decodeXmlText(relationshipTarget) : ref.location ? `#${ref.location}` : undefined
    if (!target) {
      continue
    }
    for (const address of addresses) {
      hyperlinks.push({
        sheetName,
        address,
        target,
        ...(ref.tooltip ? { tooltip: ref.tooltip } : {}),
        ...(ref.display ? { display: ref.display } : {}),
      })
    }
  }
  return hyperlinks.length > 0
    ? hyperlinks.toSorted(
        (left, right) =>
          decodeCellAddress(left.address)!.row - decodeCellAddress(right.address)!.row || left.address.localeCompare(right.address),
      )
    : undefined
}

export function readLargeSimpleSheetHyperlinkRefs(worksheetXml: string): LargeSimpleHyperlinkRef[] | null {
  if (!/<(?:[A-Za-z_][\w.-]*:)?hyperlinks\b/u.test(worksheetXml)) {
    return []
  }
  const refs: LargeSimpleHyperlinkRef[] = []
  for (const match of worksheetXml.matchAll(hyperlinkElementPattern)) {
    const tag = match[0]
    const ref = readXmlAttribute(tag, 'ref')
    if (!ref) {
      continue
    }
    if (!hyperlinkAddresses(ref)) {
      return null
    }
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    const location = readNonEmptyXmlAttribute(tag, 'location')
    const tooltip = readNonEmptyXmlAttribute(tag, 'tooltip')
    const display = readNonEmptyXmlAttribute(tag, 'display')
    refs.push({
      ref,
      ...(relationshipId ? { relationshipId } : {}),
      ...(location ? { location } : {}),
      ...(tooltip ? { tooltip } : {}),
      ...(display ? { display } : {}),
    })
  }
  return refs
}

export function readLargeSimpleSheetHyperlinkRefsFromBytes(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): LargeSimpleHyperlinkRef[] | null {
  const refs: LargeSimpleHyperlinkRef[] = []
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
    if (tagEnd === null) {
      return refs
    }
    if (tag.localName === 'hyperlink') {
      const ref = readAttributeFromTag(bytes, tag.endIndex, tagEnd, 'ref')
      if (!ref) {
        index = tagEnd + 1
        continue
      }
      if (!hyperlinkAddresses(ref)) {
        return null
      }
      const relationshipId =
        readAttributeFromTag(bytes, tag.endIndex, tagEnd, 'r:id') ?? readAttributeFromTag(bytes, tag.endIndex, tagEnd, 'id')
      const location = readNonEmptyXmlAttributeFromTag(bytes, tag.endIndex, tagEnd, 'location')
      const tooltip = readNonEmptyXmlAttributeFromTag(bytes, tag.endIndex, tagEnd, 'tooltip')
      const display = readNonEmptyXmlAttributeFromTag(bytes, tag.endIndex, tagEnd, 'display')
      refs.push({
        ref,
        ...(relationshipId ? { relationshipId } : {}),
        ...(location ? { location } : {}),
        ...(tooltip ? { tooltip } : {}),
        ...(display ? { display } : {}),
      })
    }
    index = tagEnd + 1
  }
  return refs
}

function worksheetRelationshipsPath(worksheetPath: string): string {
  const directory = worksheetPath.slice(0, worksheetPath.lastIndexOf('/'))
  const fileName = worksheetPath.slice(worksheetPath.lastIndexOf('/') + 1)
  return `${directory}/_rels/${fileName}.rels`
}

function hyperlinkAddresses(ref: string): string[] | null {
  const [startRef, endRef = startRef] = ref.split(':')
  const start = decodeCellAddress(startRef ?? '')
  const end = decodeCellAddress(endRef ?? '')
  if (!start || !end) {
    return null
  }
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const columnStart = Math.min(start.column, end.column)
  const columnEnd = Math.max(start.column, end.column)
  const cellCount = (rowEnd - rowStart + 1) * (columnEnd - columnStart + 1)
  if (cellCount > maxExpandedHyperlinkRangeCells) {
    return null
  }
  const addresses: string[] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let column = columnStart; column <= columnEnd; column += 1) {
      addresses.push(`${encodeColumnName(column)}${String(row + 1)}`)
    }
  }
  return addresses
}

function readNonEmptyXmlAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | undefined {
  const value = readAttributeFromTag(bytes, startIndex, tagEnd, attributeName)
  if (!value) {
    return undefined
  }
  const decoded = decodeXmlText(value).trim()
  return decoded.length > 0 ? decoded : undefined
}

function readNonEmptyXmlAttribute(xml: string, attributeName: string): string | undefined {
  const value = readXmlAttribute(xml, attributeName)
  if (!value) {
    return undefined
  }
  const decoded = decodeXmlText(value).trim()
  return decoded.length > 0 ? decoded : undefined
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeAscii(bytes, range.start, range.end) : null
}

function readAttributeRangeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      index += 1
      continue
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      index += 1
      continue
    }
    const valueStart = index + 1
    index = valueStart
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    const valueEnd = index
    if (attributeNameMatches(bytes, nameStart, nameEnd, attributeName)) {
      return { start: valueStart, end: valueEnd }
    }
    index += 1
  }
  return null
}

function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === slash || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === colon) {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart ? null : { localName: decodeAscii(bytes, localNameStart, index), endIndex: index }
}

function findTagEnd(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuote || byte === singleQuote) {
      quote = byte
      continue
    }
    if (byte === greaterThan) {
      return index
    }
  }
  return null
}

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  if (endIndex - startIndex !== attributeName.length) {
    return false
  }
  for (let index = 0; index < attributeName.length; index += 1) {
    if (bytes[startIndex + index] !== attributeName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function decodeAscii(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  let output = ''
  for (let index = startIndex; index < endIndex; index += 1) {
    output += String.fromCharCode(bytes[index] ?? 0)
  }
  return output
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

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/iu.exec(address.replaceAll('$', ''))
  if (!match) {
    return null
  }
  let column = 0
  for (const letter of match[1]?.toUpperCase() ?? '') {
    column = column * 26 + letter.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  if (!Number.isSafeInteger(row) || row <= 0 || column <= 0) {
    return null
  }
  return { row: row - 1, column: column - 1 }
}

function encodeColumnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    value -= 1
    output = String.fromCharCode(65 + (value % 26)) + output
    value = Math.floor(value / 26)
  }
  return output
}

function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}

function isXmlNameByte(byte: number): boolean {
  return (
    (byte >= 65 && byte <= 90) ||
    (byte >= 97 && byte <= 122) ||
    (byte >= 48 && byte <= 57) ||
    byte === 45 ||
    byte === 46 ||
    byte === colon ||
    byte === 95
  )
}
