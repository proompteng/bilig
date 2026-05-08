import type { CompiledFormula, StructuralAxisTransform } from '@bilig/formula'
import type { RuntimeFormula } from '../runtime-state.js'

export function dependencyTouchesSheet(dependency: string, sheetName: string): boolean {
  if (!dependency.includes('!')) {
    return false
  }
  const [qualifiedSheetName] = dependency.split('!')
  return qualifiedSheetName?.replace(/^'(.*)'$/, '$1') === sheetName
}

export function rangeDependencyAxisAffected(
  rangeDescriptor: { sheetId: number; row1: number; row2: number; col1: number; col2: number },
  targetSheetId: number,
  transform: StructuralAxisTransform,
): boolean {
  if (rangeDescriptor.sheetId !== targetSheetId) {
    return false
  }
  const start = transform.axis === 'row' ? rangeDescriptor.row1 : rangeDescriptor.col1
  const end = transform.axis === 'row' ? rangeDescriptor.row2 : rangeDescriptor.col2
  return !(end < transform.start || start >= transform.start + transform.count)
}

export function runtimeDirectRangeAxisAffected(
  targetSheetId: number | undefined,
  targetSheetName: string,
  transform: StructuralAxisTransform,
  range: { sheetName: string; rowStart: number; rowEnd: number; col: number } | undefined,
): boolean {
  if (!range || targetSheetId === undefined || range.sheetName !== targetSheetName) {
    return false
  }
  return rangeDependencyAxisAffected(
    {
      sheetId: targetSheetId,
      row1: range.rowStart,
      row2: range.rowEnd,
      col1: range.col,
      col2: range.col,
    },
    targetSheetId,
    transform,
  )
}

export function isStructurallyStableSimpleFormulaNode(node: CompiledFormula['optimizedAst']): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'CellRef':
      return true
    case 'ArrayConstant':
      return node.rows.every((row) => row.every(isStructurallyStableSimpleFormulaNode))
    case 'UnaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.argument)
    case 'BinaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.left) && isStructurallyStableSimpleFormulaNode(node.right)
    case 'NameRef':
    case 'OmittedArgument':
    case 'StructuredRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
    case 'CallExpr':
    case 'InvokeExpr':
      return false
  }
}

function arrayValuesEqual(left: ArrayLike<number>, right: ArrayLike<number>): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function hasStableStructuralSymbolicRangeLayout(
  current: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  next: Pick<CompiledFormula, 'symbolicRanges' | 'parsedSymbolicRanges'>,
  rangeDependencyCount: number,
): boolean {
  if (current.symbolicRanges.length !== next.symbolicRanges.length) {
    return false
  }
  if (rangeDependencyCount !== current.symbolicRanges.length) {
    return false
  }
  if (current.parsedSymbolicRanges === undefined || next.parsedSymbolicRanges === undefined) {
    return current.parsedSymbolicRanges === next.parsedSymbolicRanges
  }
  if (current.parsedSymbolicRanges.length !== next.parsedSymbolicRanges.length) {
    return false
  }
  for (let index = 0; index < current.parsedSymbolicRanges.length; index += 1) {
    if (current.parsedSymbolicRanges[index]?.refKind !== next.parsedSymbolicRanges[index]?.refKind) {
      return false
    }
  }
  return true
}

export function structuralRewritePreservesValue(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  transform: StructuralAxisTransform,
): boolean {
  return (
    transform.kind !== 'delete' &&
    rewritten.reusedProgram &&
    !rewritten.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.compiled.symbolicRanges.length === 0 &&
    formula.compiled.deps.every((dependency) => !dependency.includes(':')) &&
    formula.directLookup === undefined &&
    formula.directAggregate === undefined &&
    formula.directCriteria === undefined &&
    (formula.directScalar === undefined || transform.kind === 'insert') &&
    isStructurallyStableSimpleFormulaNode(rewritten.compiled.optimizedAst)
  )
}

export function structuralRewritePreservesBinding(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  allowsRangeReuse: boolean,
): boolean {
  return (
    rewritten.reusedProgram &&
    (formula.directAggregate !== undefined ||
      formula.directScalar !== undefined ||
      (allowsRangeReuse &&
        hasStableStructuralSymbolicRangeLayout(formula.compiled, rewritten.compiled, formula.rangeDependencies.length))) &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    rewritten.compiled.symbolicNames.length === 0 &&
    rewritten.compiled.symbolicTables.length === 0 &&
    rewritten.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directCriteria === undefined &&
    arrayValuesEqual(formula.compiled.program, rewritten.compiled.program) &&
    arrayValuesEqual(formula.compiled.constants, rewritten.compiled.constants) &&
    formula.compiled.mode === rewritten.compiled.mode &&
    (formula.directAggregate !== undefined || rewritten.compiled.symbolicRanges.length === formula.rangeDependencies.length)
  )
}

export function structuralDirectAggregateRewritePreservesValue(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  transform: StructuralAxisTransform,
): boolean {
  return (
    (transform.kind === 'insert' || transform.kind === 'move') &&
    formula.directAggregate !== undefined &&
    rewritten.reusedProgram &&
    !rewritten.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directCriteria === undefined
  )
}
