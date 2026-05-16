import type { CompiledFormula } from '@bilig/formula'
import type {
  RuntimeDirectAggregateDescriptor,
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectLookupDescriptor,
  RuntimeDirectScalarDescriptor,
  RuntimeDirectScalarOperand,
  RuntimeFormula,
} from '../runtime-state.js'

export interface PreparedFormulaBindingShape {
  readonly dependencies: {
    readonly rangeDependencies: Uint32Array
  }
  readonly compiled: CompiledFormula
  readonly directLookup: RuntimeDirectLookupDescriptor | undefined
  readonly directAggregate: RuntimeDirectAggregateDescriptor | undefined
  readonly directScalar: RuntimeDirectScalarDescriptor | undefined
  readonly directCriteria: RuntimeDirectCriteriaDescriptor | undefined
}

export function directLookupStructureEqual(
  left: RuntimeDirectLookupDescriptor | undefined,
  right: RuntimeDirectLookupDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right || left.kind !== right.kind) {
    return false
  }
  switch (left.kind) {
    case 'exact':
      return (
        right.kind === 'exact' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.prepared.sheetName === right.prepared.sheetName &&
        left.prepared.rowStart === right.prepared.rowStart &&
        left.prepared.rowEnd === right.prepared.rowEnd &&
        left.prepared.col === right.prepared.col &&
        left.searchMode === right.searchMode
      )
    case 'exact-uniform-numeric':
      return (
        right.kind === 'exact-uniform-numeric' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.sheetName === right.sheetName &&
        left.sheetId === right.sheetId &&
        left.rowStart === right.rowStart &&
        left.rowEnd === right.rowEnd &&
        left.col === right.col &&
        left.searchMode === right.searchMode
      )
    case 'approximate':
      return (
        right.kind === 'approximate' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.prepared.sheetName === right.prepared.sheetName &&
        left.prepared.rowStart === right.prepared.rowStart &&
        left.prepared.rowEnd === right.prepared.rowEnd &&
        left.prepared.col === right.prepared.col &&
        left.matchMode === right.matchMode
      )
    case 'approximate-uniform-numeric':
      return (
        right.kind === 'approximate-uniform-numeric' &&
        left.operandCellIndex === right.operandCellIndex &&
        left.sheetName === right.sheetName &&
        left.sheetId === right.sheetId &&
        left.rowStart === right.rowStart &&
        left.rowEnd === right.rowEnd &&
        left.col === right.col &&
        left.matchMode === right.matchMode
      )
  }
}

export function directCriteriaOperandEqual(
  left: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined,
  right: RuntimeDirectCriteriaDescriptor['criteriaPairs'][number]['criterion'] | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right || left.kind !== right.kind) {
    return false
  }
  if (left.kind === 'cell') {
    return right.kind === 'cell' && left.cellIndex === right.cellIndex
  }
  if (left.kind === 'cell-string-concat') {
    return (
      right.kind === 'cell-string-concat' &&
      left.cellIndex === right.cellIndex &&
      left.prefix === right.prefix &&
      left.suffix === right.suffix
    )
  }
  if (left.kind === 'cell-month-boundary-string-concat') {
    return (
      right.kind === 'cell-month-boundary-string-concat' &&
      left.cellIndex === right.cellIndex &&
      left.prefix === right.prefix &&
      left.suffix === right.suffix &&
      left.offsetMonths === right.offsetMonths
    )
  }
  return right.kind === 'literal' && JSON.stringify(left.value) === JSON.stringify(right.value)
}

function directScalarOperandEqual(left: RuntimeDirectScalarOperand | undefined, right: RuntimeDirectScalarOperand | undefined): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right || left.kind !== right.kind) {
    return false
  }
  if (left.kind === 'cell') {
    return right.kind === 'cell' && left.cellIndex === right.cellIndex
  }
  if (left.kind === 'error') {
    return right.kind === 'error' && left.code === right.code
  }
  return right.kind === 'literal-number' && Object.is(left.value, right.value)
}

function directScalarOperandCellIndex(operand: RuntimeDirectScalarOperand): number {
  return operand.kind === 'cell' ? operand.cellIndex : -1
}

export function directScalarDependencyCellsEqual(
  left: RuntimeDirectScalarDescriptor | undefined,
  right: RuntimeDirectScalarDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  const leftFirst = left.kind === 'binary' ? directScalarOperandCellIndex(left.left) : directScalarOperandCellIndex(left.operand)
  const leftSecond = left.kind === 'binary' ? directScalarOperandCellIndex(left.right) : -1
  const rightFirst = right.kind === 'binary' ? directScalarOperandCellIndex(right.left) : directScalarOperandCellIndex(right.operand)
  const rightSecond = right.kind === 'binary' ? directScalarOperandCellIndex(right.right) : -1
  if (leftSecond === -1 || rightSecond === -1) {
    return leftSecond === rightSecond && leftFirst === rightFirst
  }
  return (leftFirst === rightFirst && leftSecond === rightSecond) || (leftFirst === rightSecond && leftSecond === rightFirst)
}

function directCriteriaResultTransformsEqual(
  left: RuntimeDirectCriteriaDescriptor['resultTransforms'],
  right: RuntimeDirectCriteriaDescriptor['resultTransforms'],
): boolean {
  const leftTransforms = left ?? []
  const rightTransforms = right ?? []
  if (leftTransforms.length !== rightTransforms.length) {
    return false
  }
  for (let index = 0; index < leftTransforms.length; index += 1) {
    const leftTransform = leftTransforms[index]!
    const rightTransform = rightTransforms[index]!
    if (leftTransform.kind !== rightTransform.kind) {
      return false
    }
    if (leftTransform.kind === 'round') {
      if (rightTransform.kind !== 'round' || JSON.stringify(leftTransform.digits) !== JSON.stringify(rightTransform.digits)) {
        return false
      }
      continue
    }
    if (leftTransform.kind === 'if-empty-cell') {
      if (
        rightTransform.kind !== 'if-empty-cell' ||
        leftTransform.cellIndex !== rightTransform.cellIndex ||
        JSON.stringify(leftTransform.fallback) !== JSON.stringify(rightTransform.fallback)
      ) {
        return false
      }
      continue
    }
    if (rightTransform.kind !== 'if-error' || JSON.stringify(leftTransform.fallback) !== JSON.stringify(rightTransform.fallback)) {
      return false
    }
  }
  return true
}

export function directCriteriaStructureEqual(
  left: RuntimeDirectCriteriaDescriptor | undefined,
  right: RuntimeDirectCriteriaDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  if (left.aggregateKind !== right.aggregateKind) {
    return false
  }
  if (left.firstMatchMode !== right.firstMatchMode) {
    return false
  }
  const leftRange = left.aggregateRange
  const rightRange = right.aggregateRange
  if (
    leftRange?.regionId !== rightRange?.regionId ||
    leftRange?.sheetName !== rightRange?.sheetName ||
    leftRange?.rowStart !== rightRange?.rowStart ||
    leftRange?.rowEnd !== rightRange?.rowEnd ||
    leftRange?.col !== rightRange?.col ||
    leftRange?.length !== rightRange?.length
  ) {
    return false
  }
  if (left.criteriaPairs.length !== right.criteriaPairs.length) {
    return false
  }
  if (!directScalarOperandEqual(left.offsetOperand, right.offsetOperand)) {
    return false
  }
  if (!directCriteriaResultTransformsEqual(left.resultTransforms, right.resultTransforms)) {
    return false
  }
  for (let index = 0; index < left.criteriaPairs.length; index += 1) {
    const leftPair = left.criteriaPairs[index]!
    const rightPair = right.criteriaPairs[index]!
    if (
      leftPair.range.regionId !== rightPair.range.regionId ||
      leftPair.range.sheetName !== rightPair.range.sheetName ||
      leftPair.range.rowStart !== rightPair.range.rowStart ||
      leftPair.range.rowEnd !== rightPair.range.rowEnd ||
      leftPair.range.col !== rightPair.range.col ||
      leftPair.range.length !== rightPair.range.length ||
      !directCriteriaOperandEqual(leftPair.criterion, rightPair.criterion)
    ) {
      return false
    }
  }
  return true
}

export function directAggregateStructureEqual(
  left: RuntimeDirectAggregateDescriptor | undefined,
  right: RuntimeDirectAggregateDescriptor | undefined,
): boolean {
  if (left === right) {
    return true
  }
  if (!left || !right) {
    return false
  }
  return (
    left.aggregateKind === right.aggregateKind &&
    left.regionId === right.regionId &&
    left.sheetName === right.sheetName &&
    left.rowStart === right.rowStart &&
    left.rowEnd === right.rowEnd &&
    left.col === right.col &&
    left.colEnd === right.colEnd &&
    left.length === right.length
  )
}

export function floatArrayEqual(left: ArrayLike<number>, right: ArrayLike<number>): boolean {
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

export function uint32ArrayEqual(left: Uint32Array | readonly number[], right: Uint32Array | readonly number[]): boolean {
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

export function stringArrayEqual(left: readonly string[], right: readonly string[]): boolean {
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

export function hasStableSymbolicRangeLayout(
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

export function hasInPlaceDependencyRebindShape(existing: RuntimeFormula, prepared: PreparedFormulaBindingShape): boolean {
  return (
    uint32ArrayEqual(existing.rangeDependencies, prepared.dependencies.rangeDependencies) &&
    hasStableSymbolicRangeLayout(existing.compiled, prepared.compiled, existing.rangeDependencies.length) &&
    existing.compiled.symbolicNames.length === 0 &&
    prepared.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    prepared.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    prepared.compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    prepared.directLookup === undefined &&
    existing.directAggregate === undefined &&
    prepared.directAggregate === undefined &&
    existing.directCriteria === undefined &&
    prepared.directCriteria === undefined
  )
}

export function canRewriteCompiledPreservingBindings(existing: RuntimeFormula, compiled: CompiledFormula): boolean {
  return (
    hasStableSymbolicRangeLayout(existing.compiled, compiled, existing.rangeDependencies.length) &&
    existing.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    existing.directAggregate === undefined &&
    existing.directCriteria === undefined &&
    uint32ArrayEqual(existing.compiled.program, compiled.program) &&
    floatArrayEqual(existing.compiled.constants, compiled.constants) &&
    existing.compiled.mode === compiled.mode
  )
}

export function canRewriteCompiledPreservingDirectAggregate(existing: RuntimeFormula, compiled: CompiledFormula): boolean {
  return (
    existing.rangeDependencies.length === 0 &&
    existing.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    existing.directAggregate !== undefined &&
    existing.directCriteria === undefined &&
    existing.programLength === compiled.program.length &&
    existing.constNumberLength === compiled.constants.length &&
    existing.compiled.mode === compiled.mode
  )
}

export function canRewriteCompiledPreservingDirectScalar(existing: RuntimeFormula, compiled: CompiledFormula): boolean {
  return (
    existing.rangeDependencies.length === 0 &&
    existing.compiled.symbolicNames.length === 0 &&
    existing.compiled.symbolicTables.length === 0 &&
    existing.compiled.symbolicSpills.length === 0 &&
    compiled.symbolicNames.length === 0 &&
    compiled.symbolicTables.length === 0 &&
    compiled.symbolicSpills.length === 0 &&
    existing.directLookup === undefined &&
    existing.directAggregate === undefined &&
    existing.directScalar !== undefined &&
    existing.directCriteria === undefined &&
    existing.compiled.mode === compiled.mode
  )
}
