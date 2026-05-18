import * as XLSX from 'xlsx'

import type { WorkbookTableSnapshot } from '@bilig/protocol'

type StructuredReferenceSection = 'all' | 'data' | 'headers' | 'this-row' | 'totals'

interface StructuredReferenceParts {
  readonly section?: StructuredReferenceSection
  readonly startColumnName?: string
  readonly endColumnName?: string
}

interface StructuredReferenceRewriteContext {
  readonly formula: string
  readonly ownerSheetName: string
  readonly ownerAddress: string
  readonly tables: readonly WorkbookTableSnapshot[] | undefined
}

const namespacedSpreadsheetFormulaPattern = /^(?:msoxl|of):=/iu

export function normalizeImportedFormulaSource(formula: string): string {
  const trimmed = formula.trim()
  const prefix = namespacedSpreadsheetFormulaPattern.exec(trimmed)
  return prefix ? trimmed.slice(prefix[0].length) : formula
}

function isIdentifierStart(character: string): boolean {
  return /[A-Za-z_]/u.test(character)
}

function isIdentifierPart(character: string): boolean {
  return /[A-Za-z0-9_.]/u.test(character)
}

function skipDoubleQuotedString(source: string, startIndex: number): number {
  let index = startIndex + 1
  while (index < source.length) {
    if (source[index] === '"') {
      if (source[index + 1] === '"') {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}

function skipSingleQuotedSheetName(source: string, startIndex: number): number {
  let index = startIndex + 1
  while (index < source.length) {
    if (source[index] === "'") {
      if (source[index + 1] === "'") {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return source.length
}

function readBalancedStructuredReference(
  source: string,
  startIndex: number,
): { readonly text: string; readonly endIndex: number } | undefined {
  let depth = 0
  for (let index = startIndex; index < source.length; index += 1) {
    const character = source[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return {
          text: source.slice(startIndex + 1, index),
          endIndex: index + 1,
        }
      }
    }
  }
  return undefined
}

function normalizeStructuredReferenceSection(item: string): StructuredReferenceSection | undefined {
  const normalized = item.replace(/\s+/gu, ' ').trim().toUpperCase()
  switch (normalized) {
    case '#ALL':
      return 'all'
    case '#DATA':
      return 'data'
    case '#HEADERS':
      return 'headers'
    case '#THIS ROW':
    case '#THISROW':
    case '@':
      return 'this-row'
    case '#TOTALS':
    case '#TOTAL ROW':
    case '#TOTALS ROW':
      return 'totals'
    default:
      return undefined
  }
}

function unescapeStructuredColumnName(item: string): string {
  return item.replace(/^@/u, '').replace(/''/gu, "'").trim()
}

function hasBalancedOuterBrackets(item: string): boolean {
  if (!item.startsWith('[') || !item.endsWith(']')) {
    return false
  }
  let depth = 0
  for (let index = 0; index < item.length; index += 1) {
    const character = item[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return index === item.length - 1
      }
      if (depth < 0) {
        return false
      }
    }
  }
  return false
}

function unwrapStructuredReferenceItem(item: string): string {
  const trimmed = item.trim()
  return hasBalancedOuterBrackets(trimmed) ? trimmed.slice(1, -1).trim() : trimmed
}

function splitStructuredReferenceTopLevel(text: string, separator: ',' | ':'): string[] | undefined {
  const parts: string[] = []
  let depth = 0
  let startIndex = 0
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (character === '[') {
      depth += 1
    } else if (character === ']') {
      depth -= 1
      if (depth < 0) {
        return undefined
      }
    } else if (character === separator && depth === 0) {
      parts.push(text.slice(startIndex, index).trim())
      startIndex = index + 1
    }
  }
  if (depth !== 0) {
    return undefined
  }
  parts.push(text.slice(startIndex).trim())
  return parts
}

function parseStructuredReferenceToken(
  item: string,
): { readonly section?: StructuredReferenceSection; readonly columnName?: string } | undefined {
  let trimmed = unwrapStructuredReferenceItem(item)
  if (trimmed.length === 0) {
    return undefined
  }

  const section = normalizeStructuredReferenceSection(trimmed)
  if (section) {
    return { section }
  }

  let tokenSection: StructuredReferenceSection | undefined
  if (trimmed.startsWith('@')) {
    tokenSection = 'this-row'
    trimmed = unwrapStructuredReferenceItem(trimmed.slice(1).trim())
  }
  if (trimmed.length === 0) {
    return tokenSection ? { section: tokenSection } : undefined
  }

  return {
    ...(tokenSection ? { section: tokenSection } : {}),
    columnName: unescapeStructuredColumnName(trimmed),
  }
}

function parseStructuredReferenceParts(text: string): StructuredReferenceParts | undefined {
  if (text.trim().length === 0) {
    return {}
  }

  const items = splitStructuredReferenceTopLevel(text.trim(), ',')
  if (!items) {
    return undefined
  }

  let section: StructuredReferenceSection | undefined
  let startColumnName: string | undefined
  let endColumnName: string | undefined
  for (const item of items) {
    if (item.length === 0) {
      continue
    }

    const spanItems = splitStructuredReferenceTopLevel(item, ':')
    if (!spanItems || spanItems.length === 0 || spanItems.length > 2) {
      return undefined
    }

    if (spanItems.length === 2) {
      const spanStart = parseStructuredReferenceToken(spanItems[0] ?? '')
      const spanEnd = parseStructuredReferenceToken(spanItems[1] ?? '')
      if (!spanStart?.columnName || !spanEnd?.columnName || startColumnName !== undefined) {
        return undefined
      }
      if (spanStart.section) {
        section = spanStart.section
      }
      if (spanEnd.section && spanEnd.section !== section) {
        return undefined
      }
      startColumnName = spanStart.columnName
      endColumnName = spanEnd.columnName
      continue
    }

    const parsedItem = parseStructuredReferenceToken(item)
    if (!parsedItem) {
      continue
    }
    if (parsedItem.section) {
      section = parsedItem.section
    }
    if (parsedItem.columnName) {
      if (startColumnName !== undefined) {
        return undefined
      }
      startColumnName = parsedItem.columnName
      endColumnName = parsedItem.columnName
    }
  }

  return section || startColumnName
    ? {
        ...(section ? { section } : {}),
        ...(startColumnName ? { startColumnName } : {}),
        ...(endColumnName ? { endColumnName } : {}),
      }
    : undefined
}

function decodeAddress(address: string): XLSX.CellAddress | undefined {
  try {
    return XLSX.utils.decode_cell(address.replaceAll('$', ''))
  } catch {
    return undefined
  }
}

function encodeAddress(row: number, col: number): string {
  return XLSX.utils.encode_cell({ r: row, c: col })
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replaceAll("'", "''")}'`
}

function formatFormulaReference(sheetName: string, startRow: number, startCol: number, endRow: number, endCol: number): string {
  const startAddress = encodeAddress(startRow, startCol)
  const endAddress = encodeAddress(endRow, endCol)
  const prefix = `${quoteSheetName(sheetName)}!`
  return startAddress === endAddress ? `${prefix}${startAddress}` : `${prefix}${startAddress}:${endAddress}`
}

function normalizeStructuredColumnLookupName(columnName: string): string {
  return columnName.replace(/\s+/gu, ' ').trim().toLocaleLowerCase('en-US')
}

function findTableColumnIndex(table: WorkbookTableSnapshot, columnName: string): number {
  const normalizedColumnName = normalizeStructuredColumnLookupName(columnName)
  return table.columnNames.findIndex((candidate) => normalizeStructuredColumnLookupName(candidate) === normalizedColumnName)
}

function rewriteStructuredReference(
  table: WorkbookTableSnapshot,
  parts: StructuredReferenceParts,
  _ownerSheetName: string,
  ownerAddress: string,
): string | undefined {
  const tableStart = decodeAddress(table.startAddress)
  const tableEnd = decodeAddress(table.endAddress)
  const owner = decodeAddress(ownerAddress)
  if (!tableStart || !tableEnd || !owner) {
    return undefined
  }

  const section = parts.section ?? 'data'
  let startRow = tableStart.r + (table.headerRow && section === 'data' ? 1 : 0)
  let endRow = tableEnd.r - (table.totalsRow && section === 'data' ? 1 : 0)

  if (section === 'all') {
    startRow = tableStart.r
    endRow = tableEnd.r
  } else if (section === 'headers') {
    if (!table.headerRow) {
      return '#REF!'
    }
    startRow = tableStart.r
    endRow = tableStart.r
  } else if (section === 'totals') {
    if (!table.totalsRow) {
      return '#REF!'
    }
    startRow = tableEnd.r
    endRow = tableEnd.r
  } else if (section === 'this-row') {
    if (owner.r < tableStart.r || owner.r > tableEnd.r) {
      return '#REF!'
    }
    startRow = owner.r
    endRow = owner.r
  }

  let startCol = tableStart.c
  let endCol = tableEnd.c
  if (parts.startColumnName) {
    const startColumnIndex = findTableColumnIndex(table, parts.startColumnName)
    const endColumnIndex = findTableColumnIndex(table, parts.endColumnName ?? parts.startColumnName)
    if (startColumnIndex < 0 || endColumnIndex < 0) {
      return '#REF!'
    }
    startCol = tableStart.c + Math.min(startColumnIndex, endColumnIndex)
    endCol = tableStart.c + Math.max(startColumnIndex, endColumnIndex)
  }

  if (endRow < startRow || endCol < startCol) {
    return '#REF!'
  }
  return formatFormulaReference(table.sheetName, startRow, startCol, endRow, endCol)
}

function ownerTableForAddress(
  tables: readonly WorkbookTableSnapshot[],
  ownerSheetName: string,
  ownerAddress: string,
): WorkbookTableSnapshot | undefined {
  const owner = decodeAddress(ownerAddress)
  if (!owner) {
    return undefined
  }
  return tables.find((table) => {
    if (table.sheetName !== ownerSheetName) {
      return false
    }
    const tableStart = decodeAddress(table.startAddress)
    const tableEnd = decodeAddress(table.endAddress)
    return (
      tableStart !== undefined &&
      tableEnd !== undefined &&
      owner.r >= tableStart.r &&
      owner.r <= tableEnd.r &&
      owner.c >= tableStart.c &&
      owner.c <= tableEnd.c
    )
  })
}

function isExternalWorkbookReferencePrefix(source: string, startIndex: number): boolean {
  const match = /^\[([1-9][0-9]*)\]/u.exec(source.slice(startIndex))
  if (!match) {
    return false
  }
  let index = startIndex + match[0].length
  const sheetStart = index
  while (index < source.length && /[A-Za-z0-9_.-]/u.test(source[index] ?? '')) {
    index += 1
  }
  return index > sheetStart && source[index] === '!'
}

export function translateImportedFormulaStructuredReferences({
  formula,
  ownerSheetName,
  ownerAddress,
  tables,
}: StructuredReferenceRewriteContext): string {
  if (!tables || tables.length === 0 || !formula.includes('[')) {
    return formula
  }

  const tablesByName = new Map(tables.map((table) => [table.name.toLocaleLowerCase('en-US'), table]))
  const ownerTable = ownerTableForAddress(tables, ownerSheetName, ownerAddress)
  let output = ''
  let index = 0
  while (index < formula.length) {
    const character = formula[index]!
    if (character === '"') {
      const endIndex = skipDoubleQuotedString(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === "'") {
      const endIndex = skipSingleQuotedSheetName(formula, index)
      output += formula.slice(index, endIndex)
      index = endIndex
      continue
    }
    if (character === '[') {
      const structuredReference = !isExternalWorkbookReferencePrefix(formula, index)
        ? readBalancedStructuredReference(formula, index)
        : undefined
      const parts = structuredReference ? parseStructuredReferenceParts(structuredReference.text) : undefined
      const rewritten = ownerTable
        ? parts
          ? rewriteStructuredReference(ownerTable, parts, ownerSheetName, ownerAddress)
          : undefined
        : undefined
      if (structuredReference && rewritten) {
        output += rewritten
        index = structuredReference.endIndex
        continue
      }
      output += character
      index += 1
      continue
    }
    if (!isIdentifierStart(character)) {
      output += character
      index += 1
      continue
    }

    let identifierEnd = index + 1
    while (identifierEnd < formula.length && isIdentifierPart(formula[identifierEnd]!)) {
      identifierEnd += 1
    }
    const tableName = formula.slice(index, identifierEnd)
    const table = tablesByName.get(tableName.toLocaleLowerCase('en-US'))
    if (!table || formula[identifierEnd] !== '[') {
      output += tableName
      index = identifierEnd
      continue
    }

    const structuredReference = readBalancedStructuredReference(formula, identifierEnd)
    const parts = structuredReference ? parseStructuredReferenceParts(structuredReference.text) : undefined
    const rewritten = parts ? rewriteStructuredReference(table, parts, ownerSheetName, ownerAddress) : undefined
    if (!structuredReference || !rewritten) {
      output += tableName
      index = identifierEnd
      continue
    }
    output += rewritten
    index = structuredReference.endIndex
  }
  return output
}
