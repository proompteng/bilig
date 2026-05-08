import { parseRangeAddress, type FormulaNode } from '@bilig/formula'

export interface IndexedExactLookupCandidate {
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

export interface DirectApproximateLookupCandidate {
  sheetName?: string
  start: string
  end: string
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

export function staticIntegerValue(node: FormulaNode | undefined): number | undefined {
  if (!node) {
    return undefined
  }
  if (node.kind === 'NumberLiteral') {
    return Number.isInteger(node.value) ? node.value : undefined
  }
  if (
    node.kind === 'UnaryExpr' &&
    node.operator === '-' &&
    node.argument.kind === 'NumberLiteral' &&
    Number.isInteger(node.argument.value)
  ) {
    return -node.argument.value
  }
  return undefined
}

export function hasIndexedExactLookupCandidate(node: FormulaNode): boolean {
  return collectIndexedExactLookupCandidates(node).length > 0
}

export function hasDirectApproximateLookupCandidate(node: FormulaNode): boolean {
  return collectDirectApproximateLookupCandidates(node).length > 0
}

export function collectIndexedExactLookupCandidates(node: FormulaNode): IndexedExactLookupCandidate[] {
  switch (node.kind) {
    case 'CallExpr': {
      const callee = node.callee.trim().toUpperCase()
      const lookupRange = node.args[1]
      if (
        lookupRange?.kind === 'RangeRef' &&
        lookupRange.refKind === 'cells' &&
        lookupRange.sheetEndName === undefined &&
        lookupRange.start !== lookupRange.end
      ) {
        const isIndexedLookupCall =
          (callee === 'MATCH' && node.args.length === 3 && staticIntegerValue(node.args[2]) === 0) ||
          (callee === 'XMATCH' &&
            node.args.length >= 2 &&
            node.args.length <= 4 &&
            (node.args.length === 2 || staticIntegerValue(node.args[2]) === 0) &&
            (node.args.length < 4 || staticIntegerValue(node.args[3]) === 1 || staticIntegerValue(node.args[3]) === -1))
        if (isIndexedLookupCall) {
          const parsedRange = parseRangeAddress(`${lookupRange.start}:${lookupRange.end}`, lookupRange.sheetName)
          if (parsedRange.kind === 'cells') {
            return [
              {
                ...(lookupRange.sheetName === undefined ? {} : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectIndexedExactLookupCandidates),
            ]
          }
        }
      }
      return node.args.flatMap(collectIndexedExactLookupCandidates)
    }
    case 'UnaryExpr':
      return collectIndexedExactLookupCandidates(node.argument)
    case 'BinaryExpr':
      return [...collectIndexedExactLookupCandidates(node.left), ...collectIndexedExactLookupCandidates(node.right)]
    case 'InvokeExpr':
      return [...collectIndexedExactLookupCandidates(node.callee), ...node.args.flatMap(collectIndexedExactLookupCandidates)]
    case 'ArrayConstant':
      return node.rows.flatMap((row) => row.flatMap(collectIndexedExactLookupCandidates))
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return []
  }
}

export function collectDirectApproximateLookupCandidates(node: FormulaNode): DirectApproximateLookupCandidate[] {
  switch (node.kind) {
    case 'CallExpr': {
      const callee = node.callee.trim().toUpperCase()
      const lookupRange = node.args[1]
      if (
        lookupRange?.kind === 'RangeRef' &&
        lookupRange.refKind === 'cells' &&
        lookupRange.sheetEndName === undefined &&
        lookupRange.start !== lookupRange.end
      ) {
        const matchMode = staticIntegerValue(node.args[2])
        const searchMode = node.args.length >= 4 ? staticIntegerValue(node.args[3]) : 1
        const isDirectApproximateLookupCall =
          (callee === 'MATCH' && node.args.length === 3 && (matchMode === 1 || matchMode === -1)) ||
          (callee === 'XMATCH' &&
            node.args.length >= 3 &&
            node.args.length <= 4 &&
            (matchMode === 1 || matchMode === -1) &&
            searchMode === 1)
        if (isDirectApproximateLookupCall) {
          const parsedRange = parseRangeAddress(`${lookupRange.start}:${lookupRange.end}`, lookupRange.sheetName)
          if (parsedRange.kind === 'cells') {
            return [
              {
                ...(lookupRange.sheetName === undefined ? {} : { sheetName: lookupRange.sheetName }),
                start: lookupRange.start,
                end: lookupRange.end,
                startRow: parsedRange.start.row,
                endRow: parsedRange.end.row,
                startCol: parsedRange.start.col,
                endCol: parsedRange.end.col,
              },
              ...node.args.flatMap(collectDirectApproximateLookupCandidates),
            ]
          }
        }
      }
      return node.args.flatMap(collectDirectApproximateLookupCandidates)
    }
    case 'UnaryExpr':
      return collectDirectApproximateLookupCandidates(node.argument)
    case 'BinaryExpr':
      return [...collectDirectApproximateLookupCandidates(node.left), ...collectDirectApproximateLookupCandidates(node.right)]
    case 'InvokeExpr':
      return [...collectDirectApproximateLookupCandidates(node.callee), ...node.args.flatMap(collectDirectApproximateLookupCandidates)]
    case 'ArrayConstant':
      return node.rows.flatMap((row) => row.flatMap(collectDirectApproximateLookupCandidates))
    case 'BooleanLiteral':
    case 'CellRef':
    case 'ColumnRef':
    case 'ErrorLiteral':
    case 'NameRef':
    case 'NumberLiteral':
    case 'OmittedArgument':
    case 'RangeRef':
    case 'RowRef':
    case 'SpillRef':
    case 'StringLiteral':
    case 'StructuredRef':
      return []
  }
}
