import type { RangeAddress } from './addressing.js'
import { parseCellAddress, parseRangeAddress } from './addressing.js'
import type { ParsedCellReferenceInfo, ParsedDependencyReference, ParsedRangeReferenceInfo } from './compiler.js'
import {
  formatAxisReference,
  formatCellReference,
  parseAxisReferenceParts,
  parseCellReferenceParts,
  quoteSheetNameIfNeeded,
} from './translation-reference-utils.js'

export function translateParsedCellReference<Reference extends ParsedCellReferenceInfo>(
  reference: Reference,
  rowDelta: number,
  colDelta: number,
): Reference {
  const parts =
    reference.rowAbsolute !== undefined && reference.colAbsolute !== undefined && reference.row !== undefined && reference.col !== undefined
      ? {
          row: reference.row,
          col: reference.col,
          rowAbsolute: reference.rowAbsolute,
          colAbsolute: reference.colAbsolute,
        }
      : parseCellReferenceParts(reference.address)
  if (!parts) {
    return reference
  }
  const nextRow = parts.rowAbsolute ? parts.row : parts.row + rowDelta
  const nextCol = parts.colAbsolute ? parts.col : parts.col + colDelta
  const nextLocalAddress = formatCellReference(parts, nextRow, nextCol)
  const nextAddress =
    reference.explicitSheet || reference.sheetName !== undefined
      ? formatQualifiedCellReference(reference.sheetName, nextLocalAddress)
      : nextLocalAddress
  return {
    ...reference,
    address: nextAddress,
    ...(reference.sheetName !== undefined ? { sheetName: reference.sheetName } : {}),
    ...(reference.explicitSheet !== undefined ? { explicitSheet: reference.explicitSheet } : {}),
    ...(reference.row !== undefined ? { row: nextRow } : {}),
    ...(reference.col !== undefined ? { col: nextCol } : {}),
    ...(reference.rowAbsolute !== undefined ? { rowAbsolute: parts.rowAbsolute } : {}),
    ...(reference.colAbsolute !== undefined ? { colAbsolute: parts.colAbsolute } : {}),
  }
}

export function translateParsedRangeReference(
  reference: ParsedRangeReferenceInfo,
  rowDelta: number,
  colDelta: number,
): ParsedRangeReferenceInfo {
  const nextRange = translateParsedRangeReferenceInfo(reference, rowDelta, colDelta)
  const bounds =
    nextRange.refKind === 'cells'
      ? {
          startRow: nextRange.startRow,
          endRow: nextRange.endRow,
          startCol: nextRange.startCol,
          endCol: nextRange.endCol,
        }
      : nextRange.refKind === 'rows'
        ? {
            startRow: nextRange.startRow,
            endRow: nextRange.endRow,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: nextRange.startCol,
            endCol: nextRange.endCol,
          }
  return {
    ...reference,
    address: formatParsedRangeReference(nextRange),
    refKind: nextRange.refKind,
    startAddress: nextRange.startAddress,
    endAddress: nextRange.endAddress,
    ...bounds,
    ...(reference.explicitSheet !== undefined ? { explicitSheet: reference.explicitSheet } : {}),
  }
}

export function translateParsedDependencyReference(
  reference: ParsedDependencyReference,
  rowDelta: number,
  colDelta: number,
): ParsedDependencyReference {
  return reference.kind === 'cell'
    ? translateParsedCellReference(reference, rowDelta, colDelta)
    : translateParsedRangeReference(reference, rowDelta, colDelta)
}

export function translateQualifiedCellReference(raw: string, rowDelta: number, colDelta: number): string {
  const explicitlyQualified = raw.includes('!')
  const parsed = parseCellAddress(raw)
  const nextAddress = translateCellReference(parsed.text, rowDelta, colDelta)
  return explicitlyQualified ? formatQualifiedCellReference(parsed.sheetName, nextAddress) : nextAddress
}

export function formatParsedCellReference(reference: ParsedCellReferenceInfo): string {
  const localAddress = formatParsedLocalCellReference(reference)
  return reference.explicitSheet || reference.sheetName !== undefined
    ? formatQualifiedCellReference(reference.sheetName, localAddress)
    : localAddress
}

export function formatParsedLocalCellReference(reference: ParsedCellReferenceInfo): string {
  const parts =
    reference.row !== undefined && reference.col !== undefined && reference.rowAbsolute !== undefined && reference.colAbsolute !== undefined
      ? {
          row: reference.row,
          col: reference.col,
          rowAbsolute: reference.rowAbsolute,
          colAbsolute: reference.colAbsolute,
        }
      : parseCellReferenceParts(reference.address)
  if (!parts) {
    return stripSheetQualifier(reference.address)
  }
  return formatCellReference(parts, parts.row, parts.col)
}

export function formatParsedRangeReference(reference: ParsedRangeReferenceInfo): string {
  return formatQualifiedRangeReference(
    reference.explicitSheet ? reference.sheetName : undefined,
    reference.explicitSheet ? reference.sheetEndName : undefined,
    reference.startAddress,
    reference.endAddress,
  )
}

export function translatedCellInstructionKey(sheetName: string | undefined, address: string): string {
  return `${sheetName ?? ''}\t${address}`
}

export function translatedRangeInstructionKey(
  sheetName: string | undefined,
  sheetEndName: string | undefined,
  refKind: 'cells' | 'rows' | 'cols',
  start: string,
  end: string,
): string {
  return `${sheetName ?? ''}\t${sheetEndName ?? ''}\t${refKind}\t${start}\t${end}`
}

export function buildTranslatedCellReferenceMap(
  original: readonly ParsedCellReferenceInfo[] | undefined,
  translated: readonly ParsedCellReferenceInfo[] | undefined,
): Map<string, ParsedCellReferenceInfo> {
  const output = new Map<string, ParsedCellReferenceInfo>()
  if (!original || !translated || original.length !== translated.length) {
    return output
  }
  for (let index = 0; index < original.length; index += 1) {
    const source = original[index]
    const target = translated[index]
    if (!source || !target) {
      continue
    }
    output.set(translatedCellInstructionKey(source.sheetName, formatParsedLocalCellReference(source)), target)
  }
  return output
}

export function buildTranslatedRangeReferenceMap(
  original: readonly ParsedRangeReferenceInfo[] | undefined,
  translated: readonly ParsedRangeReferenceInfo[] | undefined,
): Map<string, ParsedRangeReferenceInfo> {
  const output = new Map<string, ParsedRangeReferenceInfo>()
  if (!original || !translated || original.length !== translated.length) {
    return output
  }
  for (let index = 0; index < original.length; index += 1) {
    const source = original[index]
    const target = translated[index]
    if (!source || !target) {
      continue
    }
    output.set(
      translatedRangeInstructionKey(source.sheetName, source.sheetEndName, source.refKind, source.startAddress, source.endAddress),
      target,
    )
  }
  return output
}

export function formatParsedDependencyReference(reference: ParsedDependencyReference): string {
  return reference.kind === 'cell' ? formatParsedCellReference(reference) : formatParsedRangeReference(reference)
}

export function translateQualifiedDependencyReference(raw: string, rowDelta: number, colDelta: number): string {
  if (!raw.includes(':')) {
    return translateQualifiedCellReference(raw, rowDelta, colDelta)
  }
  return translateQualifiedRangeReference(raw, rowDelta, colDelta)
}

export function translateQualifiedRangeReference(raw: string, rowDelta: number, colDelta: number): string {
  const parsed = parseRawQualifiedRangeReference(raw)
  const nextRange = translateRangeEndpoints(parsed.refKind, parsed.start, parsed.end, rowDelta, colDelta)
  return formatQualifiedRangeReference(parsed.sheetName, parsed.sheetEndName, nextRange.start, nextRange.end)
}

function parseRawQualifiedRangeReference(raw: string): {
  readonly sheetName?: string
  readonly sheetEndName?: string
  readonly refKind: 'cells' | 'rows' | 'cols'
  readonly start: string
  readonly end: string
} {
  const bang = raw.lastIndexOf('!')
  if (bang === -1) {
    const parsed = parseRangeAddress(raw)
    const rawRange = splitLocalRangeReference(raw)
    return {
      refKind: parsed.kind,
      start: rawRange.start,
      end: rawRange.end,
    }
  }
  const qualifier = raw.slice(0, bang).trim()
  const localRangeText = raw.slice(bang + 1).trim()
  const parsedLocalRange = parseRangeAddress(localRangeText)
  const rawRange = splitLocalRangeReference(localRangeText)
  const sheetRange = splitSheetRangeQualifier(qualifier)
  if (sheetRange) {
    return {
      sheetName: unquoteSheetQualifierPart(sheetRange.start),
      sheetEndName: unquoteSheetQualifierPart(sheetRange.end),
      refKind: parsedLocalRange.kind,
      start: rawRange.start,
      end: rawRange.end,
    }
  }
  return {
    sheetName: unquoteSheetQualifierPart(qualifier),
    refKind: parsedLocalRange.kind,
    start: rawRange.start,
    end: rawRange.end,
  }
}

function splitLocalRangeReference(raw: string): { readonly start: string; readonly end: string } {
  const separator = raw.indexOf(':')
  if (separator <= 0 || separator >= raw.length - 1) {
    throw new Error(`Invalid range address: ${raw}`)
  }
  return {
    start: raw.slice(0, separator).trim(),
    end: raw.slice(separator + 1).trim(),
  }
}

function splitSheetRangeQualifier(qualifier: string): { start: string; end: string } | undefined {
  let quoted = false
  for (let index = 0; index < qualifier.length; index += 1) {
    const char = qualifier[index]!
    if (char === "'") {
      if (quoted && qualifier[index + 1] === "'") {
        index += 1
        continue
      }
      quoted = !quoted
      continue
    }
    if (char === ':' && !quoted) {
      return {
        start: qualifier.slice(0, index),
        end: qualifier.slice(index + 1),
      }
    }
  }
  return undefined
}

function unquoteSheetQualifierPart(part: string): string {
  const trimmed = part.trim()
  return trimmed.startsWith("'") && trimmed.endsWith("'") ? trimmed.slice(1, -1).replace(/''/g, "'") : trimmed
}

function translateRangeEndpoints(
  refKind: 'cells' | 'rows' | 'cols',
  start: string,
  end: string,
  rowDelta: number,
  colDelta: number,
): { start: string; end: string } {
  if (refKind === 'cells') {
    return {
      start: translateCellReference(start, rowDelta, colDelta),
      end: translateCellReference(end, rowDelta, colDelta),
    }
  }
  if (refKind === 'rows') {
    return {
      start: translateRowReference(start, rowDelta),
      end: translateRowReference(end, rowDelta),
    }
  }
  return {
    start: translateColumnReference(start, colDelta),
    end: translateColumnReference(end, colDelta),
  }
}

export function translateRangeAddress(range: RangeAddress, rowDelta: number, colDelta: number): RangeAddress {
  switch (range.kind) {
    case 'cells': {
      const startAddress = translateCellReference(range.start.text, rowDelta, colDelta)
      const endAddress = translateCellReference(range.end.text, rowDelta, colDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, undefined, startAddress, endAddress))
    }
    case 'rows': {
      const start = translateRowReference(range.start.text, rowDelta)
      const end = translateRowReference(range.end.text, rowDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, undefined, start, end))
    }
    case 'cols': {
      const start = translateColumnReference(range.start.text, colDelta)
      const end = translateColumnReference(range.end.text, colDelta)
      return parseRangeAddress(formatQualifiedRangeReference(range.sheetName, undefined, start, end))
    }
  }
}

export function translateCellReference(ref: string, rowDelta: number, colDelta: number): string {
  const parsed = parseCellReferenceParts(ref)
  if (!parsed) {
    throw new Error(`Invalid cell reference '${ref}'`)
  }
  const nextCol = parsed.colAbsolute ? parsed.col : parsed.col + colDelta
  const nextRow = parsed.rowAbsolute ? parsed.row : parsed.row + rowDelta
  if (nextCol < 0 || nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatCellReference(parsed, nextRow, nextCol)
}

export function translateColumnReference(ref: string, colDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, 'column')
  if (!parsed) {
    throw new Error(`Invalid column reference '${ref}'`)
  }
  const nextCol = parsed.absolute ? parsed.index : parsed.index + colDelta
  if (nextCol < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatAxisReference(parsed.absolute, nextCol, 'column')
}

export function translateRowReference(ref: string, rowDelta: number): string {
  const parsed = parseAxisReferenceParts(ref, 'row')
  if (!parsed) {
    throw new Error(`Invalid row reference '${ref}'`)
  }
  const nextRow = parsed.absolute ? parsed.index : parsed.index + rowDelta
  if (nextRow < 0) {
    throw new Error(`Translated reference moved outside worksheet bounds: ${ref}`)
  }
  return formatAxisReference(parsed.absolute, nextRow, 'row')
}

function translateParsedRangeReferenceInfo(
  reference: ParsedRangeReferenceInfo,
  rowDelta: number,
  colDelta: number,
): ParsedRangeReferenceInfo {
  if (reference.refKind === 'cells') {
    const startRow = (reference.startRowAbsolute ?? false) ? reference.startRow : reference.startRow + rowDelta
    const endRow = (reference.endRowAbsolute ?? false) ? reference.endRow : reference.endRow + rowDelta
    const startCol = (reference.startColAbsolute ?? false) ? reference.startCol : reference.startCol + colDelta
    const endCol = (reference.endColAbsolute ?? false) ? reference.endCol : reference.endCol + colDelta
    const startAddress = formatCellReference(
      {
        row: reference.startRow,
        col: reference.startCol,
        rowAbsolute: reference.startRowAbsolute ?? false,
        colAbsolute: reference.startColAbsolute ?? false,
      },
      startRow,
      startCol,
    )
    const endAddress = formatCellReference(
      {
        row: reference.endRow,
        col: reference.endCol,
        rowAbsolute: reference.endRowAbsolute ?? false,
        colAbsolute: reference.endColAbsolute ?? false,
      },
      endRow,
      endCol,
    )
    return {
      ...reference,
      startAddress,
      endAddress,
      startRow,
      endRow,
      startCol,
      endCol,
    }
  }
  if (reference.refKind === 'rows') {
    const startRow = (reference.startRowAbsolute ?? false) ? reference.startRow : reference.startRow + rowDelta
    const endRow = (reference.endRowAbsolute ?? false) ? reference.endRow : reference.endRow + rowDelta
    return {
      ...reference,
      startAddress: formatAxisReference(reference.startRowAbsolute ?? false, startRow, 'row'),
      endAddress: formatAxisReference(reference.endRowAbsolute ?? false, endRow, 'row'),
      startRow,
      endRow,
      startCol: 0,
      endCol: 0,
    }
  }
  const startCol = (reference.startColAbsolute ?? false) ? reference.startCol : reference.startCol + colDelta
  const endCol = (reference.endColAbsolute ?? false) ? reference.endCol : reference.endCol + colDelta
  return {
    ...reference,
    startAddress: formatAxisReference(reference.startColAbsolute ?? false, startCol, 'column'),
    endAddress: formatAxisReference(reference.endColAbsolute ?? false, endCol, 'column'),
    startRow: 0,
    endRow: 0,
    startCol,
    endCol,
  }
}

function stripSheetQualifier(reference: string): string {
  const bang = reference.lastIndexOf('!')
  return bang === -1 ? reference : reference.slice(bang + 1)
}

function formatQualifiedCellReference(sheetName: string | undefined, address: string): string {
  if (!sheetName) {
    return address
  }
  const parsed = parseCellAddress(address, sheetName)
  return `${quoteSheetNameIfNeeded(sheetName)}!${parsed.text}`
}

function formatQualifiedRangeReference(
  sheetName: string | undefined,
  sheetEndName: string | undefined,
  start: string,
  end: string,
): string {
  const prefix =
    sheetName && sheetEndName
      ? `${quoteSheetNameIfNeeded(sheetName)}:${quoteSheetNameIfNeeded(sheetEndName)}!`
      : sheetName
        ? `${quoteSheetNameIfNeeded(sheetName)}!`
        : ''
  return `${prefix}${start}:${end}`
}
