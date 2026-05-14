import * as XLSX from 'xlsx'

import { escapeXmlAttribute } from './xlsx-export-xml.js'

interface MissingCellXml {
  readonly address: string
  readonly xml: string
}

const worksheetRowElementPattern = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/gu
const worksheetRowOpeningTagPattern = /^<row\b[^>]*(?:\/>|>)/u

function rowNumberForAddress(address: string): number {
  return XLSX.utils.decode_cell(address).r + 1
}

function columnIndexForAddress(address: string): number {
  return XLSX.utils.decode_cell(address).c
}

function rowCellXml(rowCells: readonly MissingCellXml[]): string[] {
  return rowCells
    .toSorted((left, right) => columnIndexForAddress(left.address) - columnIndexForAddress(right.address))
    .map((cell) => cell.xml)
}

function buildRowsXml(entries: readonly [number, MissingCellXml[]][]): string {
  return entries.map(([rowNumber, rowCells]) => `<row r="${String(rowNumber)}">${rowCellXml(rowCells).join('')}</row>`).join('')
}

function rowNumberForRowXml(rowXml: string): number | null {
  const rowTag = worksheetRowOpeningTagPattern.exec(rowXml)?.[0]
  const rowNumber = rowTag ? Number(/\br="([0-9]+)"/u.exec(rowTag)?.[1] ?? NaN) : NaN
  return Number.isSafeInteger(rowNumber) && rowNumber > 0 ? rowNumber : null
}

function appendCellsToRowXml(rowXml: string, cells: readonly MissingCellXml[]): string {
  const cellsXml = rowCellXml(cells).join('')
  return rowXml.endsWith('/>') ? rowXml.replace(/\/>$/u, `>${cellsXml}</row>`) : rowXml.replace('</row>', `${cellsXml}</row>`)
}

function addMissingCellsToSheetDataXml(sheetDataXml: string, rowEntries: readonly [number, MissingCellXml[]][]): string {
  const sheetDataOpeningTag = /^<sheetData\b[^>]*>/u.exec(sheetDataXml)?.[0]
  if (!sheetDataOpeningTag || !sheetDataXml.endsWith('</sheetData>')) {
    return sheetDataXml
  }

  const bodyStart = sheetDataOpeningTag.length
  const bodyEnd = sheetDataXml.length - '</sheetData>'.length
  const sheetDataBody = sheetDataXml.slice(bodyStart, bodyEnd)
  let outputBody = ''
  let lastIndex = 0
  let missingIndex = 0

  for (const match of sheetDataBody.matchAll(worksheetRowElementPattern)) {
    const rowXml = match[0]
    const rowNumber = rowNumberForRowXml(rowXml)
    outputBody += sheetDataBody.slice(lastIndex, match.index)
    if (rowNumber === null) {
      outputBody += rowXml
      lastIndex = match.index + rowXml.length
      continue
    }
    while (missingIndex < rowEntries.length && rowEntries[missingIndex]![0] < rowNumber) {
      outputBody += buildRowsXml([rowEntries[missingIndex]!])
      missingIndex += 1
    }
    if (missingIndex < rowEntries.length && rowEntries[missingIndex]![0] === rowNumber) {
      outputBody += appendCellsToRowXml(rowXml, rowEntries[missingIndex]![1])
      missingIndex += 1
    } else {
      outputBody += rowXml
    }
    lastIndex = match.index + rowXml.length
  }

  outputBody += sheetDataBody.slice(lastIndex)
  while (missingIndex < rowEntries.length) {
    outputBody += buildRowsXml([rowEntries[missingIndex]!])
    missingIndex += 1
  }

  return `${sheetDataOpeningTag}${outputBody}</sheetData>`
}

export function addMissingCellsToSheetXml(sheetXml: string, cells: readonly MissingCellXml[]): string {
  if (cells.length === 0) {
    return sheetXml
  }
  const byRow = new Map<number, MissingCellXml[]>()
  for (const cell of cells) {
    const rowNumber = rowNumberForAddress(cell.address)
    byRow.set(rowNumber, [...(byRow.get(rowNumber) ?? []), cell])
  }
  const rowEntries = [...byRow.entries()].toSorted(([left], [right]) => left - right)
  const selfClosingSheetData = /<sheetData\b([^>]*)\/>/u
  if (selfClosingSheetData.test(sheetXml)) {
    return sheetXml.replace(
      selfClosingSheetData,
      (_match, attributes: string) => `<sheetData${attributes}>${buildRowsXml(rowEntries)}</sheetData>`,
    )
  }

  const sheetDataMatch = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/u.exec(sheetXml)
  if (sheetDataMatch) {
    return `${sheetXml.slice(0, sheetDataMatch.index)}${addMissingCellsToSheetDataXml(sheetDataMatch[0], rowEntries)}${sheetXml.slice(
      sheetDataMatch.index + sheetDataMatch[0].length,
    )}`
  }

  return sheetXml.replace('</worksheet>', `<sheetData>${buildRowsXml(rowEntries)}</sheetData></worksheet>`)
}

export function addMissingFormattedCells(
  sheetXml: string,
  cells: readonly { readonly address: string; readonly styleIndex: number }[],
): string {
  return addMissingCellsToSheetXml(
    sheetXml,
    cells.map((cell) => {
      const address = escapeXmlAttribute(cell.address)
      return {
        address: cell.address,
        xml: `<c r="${address}" s="${String(cell.styleIndex)}"/>`,
      }
    }),
  )
}

export function addMissingBlankCells(sheetXml: string, addresses: readonly string[]): string {
  return addMissingCellsToSheetXml(
    sheetXml,
    addresses.map((address) => ({
      address,
      xml: `<c r="${escapeXmlAttribute(address)}"/>`,
    })),
  )
}
