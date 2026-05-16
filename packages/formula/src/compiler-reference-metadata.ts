import type { FormulaNode } from './ast.js'
import { formatRangeAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import type { ParsedCellReferenceInfo, ParsedDependencyReference, ParsedRangeReferenceInfo } from './compiler-types.js'
import { rewriteSpecialCall } from './special-call-rewrites.js'
import {
  parseAxisReferenceParts as parseLocalAxisReferenceParts,
  parseCellReferenceParts as parseLocalCellReferenceParts,
  quoteSheetNameIfNeeded,
} from './translation-reference-utils.js'

function stripSheetQualifier(reference: string): string {
  const bang = reference.lastIndexOf('!')
  return bang === -1 ? reference.trim() : reference.slice(bang + 1).trim()
}

function parseCellReferenceParts(reference: string): ReturnType<typeof parseLocalCellReferenceParts> {
  return parseLocalCellReferenceParts(stripSheetQualifier(reference))
}

function parseAxisReferenceParts(reference: string, kind: 'row' | 'column'): ReturnType<typeof parseLocalAxisReferenceParts> {
  return parseLocalAxisReferenceParts(stripSheetQualifier(reference), kind)
}

export function formatRangeReference(sheetName: string | undefined, sheetEndName: string | undefined, start: string, end: string): string {
  if (sheetEndName !== undefined) {
    if (sheetName === undefined) {
      throw new Error('Sheet range references require a start sheet')
    }
    return `${quoteSheetNameIfNeeded(sheetName)}:${quoteSheetNameIfNeeded(sheetEndName)}!${start}:${end}`
  }
  return sheetName ? `${quoteSheetNameIfNeeded(sheetName)}!${start}:${end}` : `${start}:${end}`
}

export function formatRangeReferenceNode(node: Extract<FormulaNode, { kind: 'RangeRef' }>): string {
  return formatRangeReference(node.sheetName, node.sheetEndName, node.start, node.end)
}

export function buildParsedCellReferenceInfo(reference: string): ParsedCellReferenceInfo {
  const parsedCell = parseCellAddress(reference)
  const parts = parseCellReferenceParts(reference)
  return {
    address: reference,
    ...(parsedCell.sheetName !== undefined ? { sheetName: parsedCell.sheetName } : {}),
    ...(reference.includes('!') ? { explicitSheet: true } : {}),
    row: parsedCell.row,
    col: parsedCell.col,
    ...(parts
      ? {
          rowAbsolute: parts.rowAbsolute,
          colAbsolute: parts.colAbsolute,
        }
      : {}),
  }
}

function buildParsedRangeReferenceInfo(reference: string): ParsedRangeReferenceInfo {
  const parsedRange = parseRangeAddress(reference)
  const bounds =
    parsedRange.kind === 'cells'
      ? {
          startRow: parsedRange.start.row,
          endRow: parsedRange.end.row,
          startCol: parsedRange.start.col,
          endCol: parsedRange.end.col,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRow: parsedRange.start.row,
            endRow: parsedRange.end.row,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: parsedRange.start.col,
            endCol: parsedRange.end.col,
          }
  const separator = reference.indexOf(':')
  const rawStart = reference.slice(0, separator).trim()
  const rawEnd = reference.slice(separator + 1).trim()
  const cellStart = parsedRange.kind === 'cells' ? parseCellReferenceParts(rawStart) : undefined
  const cellEnd = parsedRange.kind === 'cells' ? parseCellReferenceParts(rawEnd) : undefined
  const rowStart = parsedRange.kind === 'rows' ? parseAxisReferenceParts(rawStart, 'row') : undefined
  const rowEnd = parsedRange.kind === 'rows' ? parseAxisReferenceParts(rawEnd, 'row') : undefined
  const colStart = parsedRange.kind === 'cols' ? parseAxisReferenceParts(rawStart, 'column') : undefined
  const colEnd = parsedRange.kind === 'cols' ? parseAxisReferenceParts(rawEnd, 'column') : undefined
  return {
    address: reference,
    kind: 'range',
    refKind: parsedRange.kind,
    ...(parsedRange.sheetName !== undefined ? { sheetName: parsedRange.sheetName } : {}),
    ...(rawStart.includes('!') ? { explicitSheet: true } : {}),
    startAddress: parsedRange.start.text,
    endAddress: parsedRange.end.text,
    ...bounds,
    ...(parsedRange.kind === 'cells'
      ? {
          startRowAbsolute: cellStart?.rowAbsolute ?? false,
          endRowAbsolute: cellEnd?.rowAbsolute ?? false,
          startColAbsolute: cellStart?.colAbsolute ?? false,
          endColAbsolute: cellEnd?.colAbsolute ?? false,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRowAbsolute: rowStart?.absolute ?? false,
            endRowAbsolute: rowEnd?.absolute ?? false,
          }
        : {
            startColAbsolute: colStart?.absolute ?? false,
            endColAbsolute: colEnd?.absolute ?? false,
          }),
  }
}

function buildParsedRangeReferenceInfoFromNode(node: Extract<FormulaNode, { kind: 'RangeRef' }>): ParsedRangeReferenceInfo {
  if (node.sheetEndName === undefined) {
    return buildParsedRangeReferenceInfo(formatRangeReferenceNode(node))
  }
  const parsedRange = parseRangeAddress(`${node.start}:${node.end}`)
  const bounds =
    parsedRange.kind === 'cells'
      ? {
          startRow: parsedRange.start.row,
          endRow: parsedRange.end.row,
          startCol: parsedRange.start.col,
          endCol: parsedRange.end.col,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRow: parsedRange.start.row,
            endRow: parsedRange.end.row,
            startCol: 0,
            endCol: 0,
          }
        : {
            startRow: 0,
            endRow: 0,
            startCol: parsedRange.start.col,
            endCol: parsedRange.end.col,
          }
  const cellStart = parsedRange.kind === 'cells' ? parseCellReferenceParts(node.start) : undefined
  const cellEnd = parsedRange.kind === 'cells' ? parseCellReferenceParts(node.end) : undefined
  const rowStart = parsedRange.kind === 'rows' ? parseAxisReferenceParts(node.start, 'row') : undefined
  const rowEnd = parsedRange.kind === 'rows' ? parseAxisReferenceParts(node.end, 'row') : undefined
  const colStart = parsedRange.kind === 'cols' ? parseAxisReferenceParts(node.start, 'column') : undefined
  const colEnd = parsedRange.kind === 'cols' ? parseAxisReferenceParts(node.end, 'column') : undefined
  return {
    address: formatRangeReferenceNode(node),
    kind: 'range',
    refKind: parsedRange.kind,
    ...(node.sheetName === undefined ? {} : { sheetName: node.sheetName }),
    sheetEndName: node.sheetEndName,
    explicitSheet: true,
    startAddress: parsedRange.start.text,
    endAddress: parsedRange.end.text,
    ...bounds,
    ...(parsedRange.kind === 'cells'
      ? {
          startRowAbsolute: cellStart?.rowAbsolute ?? false,
          endRowAbsolute: cellEnd?.rowAbsolute ?? false,
          startColAbsolute: cellStart?.colAbsolute ?? false,
          endColAbsolute: cellEnd?.colAbsolute ?? false,
        }
      : parsedRange.kind === 'rows'
        ? {
            startRowAbsolute: rowStart?.absolute ?? false,
            endRowAbsolute: rowEnd?.absolute ?? false,
          }
        : {
            startColAbsolute: colStart?.absolute ?? false,
            endColAbsolute: colEnd?.absolute ?? false,
          }),
  }
}

function parseDependencyReference(reference: string): ParsedDependencyReference {
  if (reference.includes(':')) {
    return buildParsedRangeReferenceInfo(reference)
  }
  return {
    kind: 'cell',
    ...buildParsedCellReferenceInfo(reference),
  }
}

function normalizedDependencyReferenceKey(reference: string): string | undefined {
  try {
    if (reference.includes(':')) {
      return formatRangeAddress(parseRangeAddress(reference))
    }
    const parsedCell = parseCellAddress(reference)
    return parsedCell.sheetName === undefined ? parsedCell.text : `${parsedCell.sheetName}!${parsedCell.text}`
  } catch {
    return undefined
  }
}

function registerParsedDependencyReference(
  referencesByKey: Map<string, ParsedDependencyReference>,
  reference: string,
  parsed: ParsedDependencyReference,
): void {
  if (!referencesByKey.has(reference)) {
    referencesByKey.set(reference, parsed)
  }
  const normalized = normalizedDependencyReferenceKey(reference)
  if (normalized !== undefined && !referencesByKey.has(normalized)) {
    referencesByKey.set(normalized, parsed)
  }
}

export function collectParsedDependencyReferencesFromAst(ast: FormulaNode): Map<string, ParsedDependencyReference> {
  const referencesByKey = new Map<string, ParsedDependencyReference>()

  const collect = (node: FormulaNode): void => {
    switch (node.kind) {
      case 'NumberLiteral':
      case 'BooleanLiteral':
      case 'StringLiteral':
      case 'ErrorLiteral':
      case 'OmittedArgument':
      case 'NameRef':
      case 'StructuredRef':
      case 'SpillRef':
      case 'RowRef':
      case 'ColumnRef':
        return
      case 'ArrayConstant':
        node.rows.forEach((row) => row.forEach(collect))
        return
      case 'CellRef': {
        const reference = node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref
        registerParsedDependencyReference(referencesByKey, reference, {
          kind: 'cell',
          ...buildParsedCellReferenceInfo(reference),
        })
        return
      }
      case 'RangeRef': {
        const reference = formatRangeReferenceNode(node)
        registerParsedDependencyReference(referencesByKey, reference, buildParsedRangeReferenceInfoFromNode(node))
        return
      }
      case 'UnaryExpr':
        collect(node.argument)
        return
      case 'BinaryExpr':
        collect(node.left)
        collect(node.right)
        return
      case 'CallExpr': {
        const rewritten = rewriteSpecialCall(node)
        if (rewritten) {
          collect(rewritten)
          return
        }
        node.args.forEach(collect)
        return
      }
      case 'InvokeExpr':
        collect(node.callee)
        node.args.forEach(collect)
        return
    }
  }

  collect(ast)
  return referencesByKey
}

function resolveParsedDependencyReference(
  referencesByKey: ReadonlyMap<string, ParsedDependencyReference>,
  reference: string,
): ParsedDependencyReference | undefined {
  const normalized = normalizedDependencyReferenceKey(reference)
  return referencesByKey.get(reference) ?? (normalized === undefined ? undefined : referencesByKey.get(normalized))
}

export function buildParsedDependenciesFromReferences(
  referencesByKey: ReadonlyMap<string, ParsedDependencyReference>,
  deps: readonly string[],
): ParsedDependencyReference[] {
  return deps.map((dependency) => resolveParsedDependencyReference(referencesByKey, dependency) ?? parseDependencyReference(dependency))
}

export function buildParsedSymbolicRangesFromReferences(
  referencesByKey: ReadonlyMap<string, ParsedDependencyReference>,
  ranges: readonly string[],
): ParsedRangeReferenceInfo[] {
  return ranges.map((reference) => {
    const parsed = resolveParsedDependencyReference(referencesByKey, reference)
    return parsed?.kind === 'range' ? parsed : buildParsedRangeReferenceInfo(reference)
  })
}
