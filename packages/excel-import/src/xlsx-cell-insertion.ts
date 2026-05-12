import * as XLSX from 'xlsx'

import { escapeXmlAttribute } from './xlsx-export-xml.js'

interface MissingCellXml {
  readonly address: string
  readonly xml: string
}

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

export function addMissingCellsToSheetXml(sheetXml: string, cells: readonly MissingCellXml[]): string {
  let output = sheetXml
  const byRow = new Map<number, MissingCellXml[]>()
  for (const cell of cells) {
    const rowNumber = rowNumberForAddress(cell.address)
    byRow.set(rowNumber, [...(byRow.get(rowNumber) ?? []), cell])
  }
  const rowEntries = [...byRow.entries()].toSorted(([left], [right]) => left - right)
  const selfClosingSheetData = /<sheetData\b([^>]*)\/>/u
  if (selfClosingSheetData.test(output)) {
    return output.replace(
      selfClosingSheetData,
      (_match, attributes: string) => `<sheetData${attributes}>${buildRowsXml(rowEntries)}</sheetData>`,
    )
  }
  for (const [rowNumber, rowCells] of rowEntries) {
    const cellsXml = rowCellXml(rowCells).join('')
    const rowPattern = new RegExp(`<row\\b(?=[^>]*\\br="${String(rowNumber)}")[^>]*(?:/>|>[\\s\\S]*?</row>)`, 'u')
    if (rowPattern.test(output)) {
      output = output.replace(rowPattern, (rowXml) =>
        rowXml.endsWith('/>') ? rowXml.replace(/\/>$/u, `>${cellsXml}</row>`) : rowXml.replace('</row>', `${cellsXml}</row>`),
      )
    } else {
      output = output.replace('</sheetData>', `<row r="${String(rowNumber)}">${cellsXml}</row></sheetData>`)
    }
  }
  if (!/<sheetData\b/u.test(output)) {
    output = output.replace('</worksheet>', `<sheetData>${buildRowsXml(rowEntries)}</sheetData></worksheet>`)
  }
  return output
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
        xml: `<c r="${address}" s="${String(cell.styleIndex)}" t="z"></c>`,
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
