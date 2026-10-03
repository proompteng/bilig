import type { WorkPaperCellAddress } from './work-paper.js'

interface XlsxRecalcTargetWorkbook {
  getSheetId(sheetName: string): number | undefined
}

export function parseQualifiedCellTarget(workbook: XlsxRecalcTargetWorkbook, target: string): WorkPaperCellAddress {
  const parsed = parseQualifiedA1(target)
  const sheet = workbook.getSheetId(parsed.sheetName)
  if (sheet === undefined) {
    throw new Error(`Unknown sheet in XLSX formula recalculation target: ${parsed.sheetName}`)
  }
  return {
    sheet,
    row: parsed.row,
    col: parsed.col,
  }
}

export function parseQualifiedA1(target: string): { sheetName: string; row: number; col: number } {
  const trimmed = target.trim()
  const separator = findSheetSeparator(trimmed)
  if (separator <= 0 || separator >= trimmed.length - 1) {
    throw new Error(`Expected a sheet-qualified A1 target such as Inputs!B2, received: ${target}`)
  }

  const sheetName = unquoteSheetName(trimmed.slice(0, separator))
  const a1 = trimmed
    .slice(separator + 1)
    .replace(/\$/gu, '')
    .toUpperCase()
  const match = /^(?<col>[A-Z]+)(?<row>[1-9][0-9]*)$/u.exec(a1)
  if (!match?.groups) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  const row = match.groups['row']
  const col = match.groups['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference in target ${target}`)
  }

  return {
    sheetName,
    ...parseA1Parts(row, col),
  }
}

export function parseA1CellReference(address: string): { row: number; col: number } {
  const match = /^\$?(?<col>[A-Z]+)\$?(?<row>[1-9][0-9]*)$/u.exec(address.trim().toUpperCase())
  if (!match?.groups) {
    throw new Error(`Expected a single A1 cell reference, received: ${address}`)
  }

  const row = match.groups['row']
  const col = match.groups['col']
  if (!row || !col) {
    throw new Error(`Expected a single A1 cell reference, received: ${address}`)
  }

  return parseA1Parts(row, col)
}

function parseA1Parts(row: string, col: string): { row: number; col: number } {
  return {
    row: Number.parseInt(row, 10) - 1,
    col: columnLettersToIndex(col),
  }
}

function findSheetSeparator(target: string): number {
  let inQuote = false
  for (let index = 0; index < target.length; index += 1) {
    const char = target[index]
    if (char === "'") {
      if (inQuote && target[index + 1] === "'") {
        index += 1
      } else {
        inQuote = !inQuote
      }
      continue
    }
    if (char === '!' && !inQuote) {
      return index
    }
  }
  return -1
}

function unquoteSheetName(rawSheetName: string): string {
  const trimmed = rawSheetName.trim()
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    return trimmed.slice(1, -1).replace(/''/gu, "'")
  }
  return trimmed
}

function columnLettersToIndex(letters: string): number {
  let index = 0
  for (const char of letters) {
    index = index * 26 + (char.charCodeAt(0) - 64)
  }
  return index - 1
}
