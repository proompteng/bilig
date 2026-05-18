import type { WorkbookHyperlinkSnapshot } from '@bilig/protocol'
import { getZipText, type XlsxZipEntries } from './xlsx-zip.js'
import { parseRelationships } from './xlsx-pivot-artifacts.js'

const hyperlinkRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
const hyperlinkElementPattern = /<(?:[A-Za-z_][\w.-]*:)?hyperlink\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu
const maxExpandedHyperlinkRangeCells = 1_024

export function readLargeSimpleSheetHyperlinks(
  zip: XlsxZipEntries,
  sheetName: string,
  worksheetPath: string,
  worksheetXml: string,
): WorkbookHyperlinkSnapshot[] | null | undefined {
  if (!/<(?:[A-Za-z_][\w.-]*:)?hyperlinks\b/u.test(worksheetXml)) {
    return undefined
  }
  const relationships = new Map(
    parseRelationships(getZipText(zip, worksheetRelationshipsPath(worksheetPath)))
      .filter((relationship) => relationship.type === hyperlinkRelationshipType || relationship.type.endsWith('/hyperlink'))
      .map((relationship) => [relationship.id, relationship]),
  )
  const hyperlinks: WorkbookHyperlinkSnapshot[] = []
  for (const match of worksheetXml.matchAll(hyperlinkElementPattern)) {
    const tag = match[0]
    const ref = readXmlAttribute(tag, 'ref')
    if (!ref) {
      continue
    }
    const addresses = hyperlinkAddresses(ref)
    if (!addresses) {
      return null
    }
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    const relationshipTarget = relationshipId ? relationships.get(relationshipId)?.target : undefined
    const location = readXmlAttribute(tag, 'location')
    const target = relationshipTarget ?? (location ? `#${decodeXmlText(location)}` : undefined)
    if (!target) {
      continue
    }
    const tooltip = readNonEmptyXmlAttribute(tag, 'tooltip')
    const display = readNonEmptyXmlAttribute(tag, 'display')
    for (const address of addresses) {
      hyperlinks.push({
        sheetName,
        address,
        target: decodeXmlText(target),
        ...(tooltip ? { tooltip } : {}),
        ...(display ? { display } : {}),
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
