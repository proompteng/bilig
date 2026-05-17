import type { FormulaNode } from './ast.js'
import { formatAddress, formatRangeAddress, parseCellAddress, parseRangeAddress } from './addressing.js'
import { hasBuiltin } from './builtins.js'
import { rewriteSpecialCall } from './special-call-rewrites.js'
import { quoteSheetNameIfNeeded } from './translation-reference-utils.js'

const MAX_EXPANDED_OFFSET_DEPENDENCY_CELLS = 4096

interface OffsetReferenceBounds {
  readonly sheetName?: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
}

export interface FormulaDependencyMetadata {
  readonly deps: string[]
  readonly symbolicNames: string[]
  readonly symbolicTables: string[]
  readonly symbolicSpills: string[]
}

function assertNever(value: never): never {
  throw new Error(`Unexpected formula node: ${JSON.stringify(value)}`)
}

function formatCellAsSingleCellRange(ref: string, sheetName: string | undefined): string {
  return formatRangeAddress(parseRangeAddress(sheetName ? `${sheetName}!${ref}:${ref}` : `${ref}:${ref}`))
}

function formatRangeDependency(node: Extract<FormulaNode, { kind: 'RangeRef' }>): string {
  if (node.sheetEndName !== undefined) {
    if (node.sheetName === undefined) {
      throw new Error('Sheet range references require a start sheet')
    }
    return `${quoteSheetNameIfNeeded(node.sheetName)}:${quoteSheetNameIfNeeded(node.sheetEndName)}!${node.start}:${node.end}`
  }
  return formatRangeAddress(parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`))
}

function staticIntegerValue(node: FormulaNode | undefined): number | undefined {
  if (!node) {
    return undefined
  }
  if (node.kind === 'NumberLiteral' && Number.isInteger(node.value)) {
    return node.value
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

function offsetReferenceBounds(node: FormulaNode | undefined): OffsetReferenceBounds | undefined {
  if (!node) {
    return undefined
  }
  if (node.kind === 'CellRef') {
    const address = parseCellAddress(node.ref, node.sheetName)
    return {
      ...(address.sheetName === undefined ? {} : { sheetName: address.sheetName }),
      rowStart: address.row,
      rowEnd: address.row,
      colStart: address.col,
      colEnd: address.col,
    }
  }
  if (node.kind !== 'RangeRef' || node.refKind !== 'cells' || node.sheetEndName !== undefined) {
    return undefined
  }
  const range = parseRangeAddress(node.sheetName ? `${node.sheetName}!${node.start}:${node.end}` : `${node.start}:${node.end}`)
  if (range.kind !== 'cells') {
    return undefined
  }
  return {
    ...(range.sheetName === undefined ? {} : { sheetName: range.sheetName }),
    rowStart: Math.min(range.start.row, range.end.row),
    rowEnd: Math.max(range.start.row, range.end.row),
    colStart: Math.min(range.start.col, range.end.col),
    colEnd: Math.max(range.start.col, range.end.col),
  }
}

function matchResultBounds(node: FormulaNode): { readonly min: number; readonly max: number } | undefined {
  if (node.kind !== 'CallExpr') {
    return undefined
  }
  const callee = node.callee.toUpperCase()
  if (callee !== 'MATCH' && callee !== 'XMATCH') {
    return undefined
  }
  const lookupRange = node.args[1]
  if (!lookupRange || lookupRange.kind !== 'RangeRef' || lookupRange.refKind !== 'cells') {
    return undefined
  }
  const range = parseRangeAddress(
    lookupRange.sheetName ? `${lookupRange.sheetName}!${lookupRange.start}:${lookupRange.end}` : `${lookupRange.start}:${lookupRange.end}`,
  )
  if (range.kind !== 'cells') {
    return undefined
  }
  const rowCount = Math.abs(range.end.row - range.start.row) + 1
  const colCount = Math.abs(range.end.col - range.start.col) + 1
  if (rowCount !== 1 && colCount !== 1) {
    return undefined
  }
  return {
    min: 1,
    max: Math.max(rowCount, colCount),
  }
}

function offsetArgumentBounds(node: FormulaNode | undefined): { readonly min: number; readonly max: number } | undefined {
  const staticValue = staticIntegerValue(node)
  if (staticValue !== undefined) {
    return { min: staticValue, max: staticValue }
  }
  return node ? matchResultBounds(node) : undefined
}

function offsetSize(node: FormulaNode | undefined, fallback: number): number | undefined {
  const value = staticIntegerValue(node)
  return value === undefined ? fallback : value
}

function qualifiedCellDependency(sheetName: string | undefined, row: number, col: number): string {
  const address = formatAddress(row, col)
  return sheetName === undefined ? address : `${sheetName}!${address}`
}

function collectOffsetTargetDependencies(node: Extract<FormulaNode, { kind: 'CallExpr' }>): string[] {
  if (node.args.length < 3 || node.args.length > 5) {
    return []
  }
  const reference = offsetReferenceBounds(node.args[0])
  const rowOffset = offsetArgumentBounds(node.args[1])
  const colOffset = offsetArgumentBounds(node.args[2])
  if (!reference || !rowOffset || !colOffset) {
    return []
  }
  const referenceRows = reference.rowEnd - reference.rowStart + 1
  const referenceCols = reference.colEnd - reference.colStart + 1
  const height = offsetSize(node.args[3], referenceRows)
  const width = offsetSize(node.args[4], referenceCols)
  if (height === undefined || width === undefined || height < 1 || width < 1) {
    return []
  }
  const rowStart = reference.rowStart + rowOffset.min
  const rowEnd = reference.rowStart + rowOffset.max + height - 1
  const colStart = reference.colStart + colOffset.min
  const colEnd = reference.colStart + colOffset.max + width - 1
  const cellCount = (rowEnd - rowStart + 1) * (colEnd - colStart + 1)
  if (rowStart < 0 || colStart < 0 || rowEnd < rowStart || colEnd < colStart || cellCount > MAX_EXPANDED_OFFSET_DEPENDENCY_CELLS) {
    return []
  }

  const dependencies: string[] = []
  for (let row = rowStart; row <= rowEnd; row += 1) {
    for (let col = colStart; col <= colEnd; col += 1) {
      dependencies.push(qualifiedCellDependency(reference.sheetName, row, col))
    }
  }
  return dependencies
}

export function collectFormulaDependencyMetadata(ast: FormulaNode): FormulaDependencyMetadata {
  const deps = new Set<string>()
  const symbolicNames = new Set<string>()
  const symbolicTables = new Set<string>()
  const symbolicSpills = new Set<string>()

  function collectDeps(node: FormulaNode, localNames: ReadonlySet<string> = new Set()): void {
    switch (node.kind) {
      case 'NumberLiteral':
      case 'BooleanLiteral':
      case 'StringLiteral':
      case 'ErrorLiteral':
      case 'OmittedArgument':
        break
      case 'ArrayConstant':
        node.rows.forEach((row) => row.forEach((entry) => collectDeps(entry, localNames)))
        break
      case 'NameRef':
        if (!localNames.has(node.name)) {
          symbolicNames.add(node.name)
        }
        break
      case 'StructuredRef':
        symbolicTables.add(node.tableName)
        break
      case 'CellRef':
        deps.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref)
        break
      case 'SpillRef':
        symbolicSpills.add(node.sheetName ? `${node.sheetName}!${node.ref}` : node.ref)
        break
      case 'RowRef':
      case 'ColumnRef':
        throw new Error('Row and column references must appear inside a range')
      case 'RangeRef':
        deps.add(formatRangeDependency(node))
        break
      case 'UnaryExpr':
        collectDeps(node.argument, localNames)
        break
      case 'BinaryExpr':
        collectDeps(node.left, localNames)
        collectDeps(node.right, localNames)
        break
      case 'CallExpr': {
        const rewritten = rewriteSpecialCall(node)
        if (rewritten) {
          collectDeps(rewritten, localNames)
          break
        }
        const callee = node.callee.toUpperCase()
        if (callee === 'LET' && node.args.length >= 3 && node.args.length % 2 === 1) {
          const scopedNames = new Set(localNames)
          for (let index = 0; index < node.args.length - 1; index += 2) {
            const nameNode = node.args[index]!
            collectDeps(node.args[index + 1]!, scopedNames)
            if (nameNode.kind === 'NameRef') {
              scopedNames.add(nameNode.name)
            }
          }
          collectDeps(node.args[node.args.length - 1]!, scopedNames)
          break
        }
        if (callee === 'LAMBDA' && node.args.length >= 1) {
          const scopedNames = new Set(localNames)
          for (let index = 0; index < node.args.length - 1; index += 1) {
            const paramNode = node.args[index]!
            if (paramNode.kind === 'NameRef') {
              scopedNames.add(paramNode.name)
            }
          }
          collectDeps(node.args[node.args.length - 1]!, scopedNames)
          break
        }
        if (!hasBuiltin(callee) && !localNames.has(node.callee)) {
          symbolicNames.add(node.callee)
        }
        const aggregateArgumentIndex = callee === 'GROUPBY' ? 2 : callee === 'PIVOTBY' ? 3 : -1
        node.args.forEach((arg, index) => {
          if (index === aggregateArgumentIndex && arg.kind === 'NameRef') {
            return
          }
          if (callee === 'SUM' && arg.kind === 'CellRef') {
            deps.add(formatCellAsSingleCellRange(arg.ref, arg.sheetName))
            return
          }
          collectDeps(arg, localNames)
        })
        if (callee === 'OFFSET') {
          collectOffsetTargetDependencies(node).forEach((dependency) => deps.add(dependency))
        }
        break
      }
      case 'InvokeExpr':
        collectDeps(node.callee, localNames)
        node.args.forEach((arg) => {
          collectDeps(arg, localNames)
        })
        break
      default:
        assertNever(node)
    }
  }

  collectDeps(ast)
  return {
    deps: [...deps],
    symbolicNames: [...symbolicNames],
    symbolicTables: [...symbolicTables],
    symbolicSpills: [...symbolicSpills],
  }
}
