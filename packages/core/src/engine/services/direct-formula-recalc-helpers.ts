import { MAX_COLS, ValueTag, type CellValue } from '@bilig/protocol'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'
import type { RuntimeDirectCriteriaDescriptor, U32 } from '../runtime-state.js'
import type { DirectFormulaIndexCollection } from './direct-formula-index-collection.js'

export type DirectFormulaRecalcRecord = {
  readonly directAggregate?: unknown
  readonly directCriteria?: unknown
  readonly directLookup?: unknown
  readonly directScalar?: unknown
}

export function directAggregateNumericContribution(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
    case ValueTag.String:
      return 0
    case ValueTag.Error:
      return undefined
  }
}

export function directCriteriaValueString(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return ''
    case ValueTag.Number:
      return String(Object.is(value.value, -0) ? 0 : value.value)
    case ValueTag.Boolean:
      return value.value ? 'TRUE' : 'FALSE'
    case ValueTag.String:
      return value.value
    case ValueTag.Error:
      return String(value.code)
  }
}

export function collectTrackedDependents<Key extends string | number>(registry: Map<Key, Set<number>>, keys: readonly Key[]): number[] {
  const candidates = new Set<number>()
  keys.forEach((key) => {
    registry.get(key)?.forEach((cellIndex) => {
      candidates.add(cellIndex)
    })
  })
  return [...candidates]
}

export function directCriteriaTouchesPoint(
  directCriteria: RuntimeDirectCriteriaDescriptor,
  request: { sheetName: string; row: number; col: number; inputCellIndex?: number },
): boolean {
  if (directCriteria.aggregateRange) {
    const aggregateRange = directCriteria.aggregateRange
    if (
      aggregateRange.sheetName === request.sheetName &&
      aggregateRange.col === request.col &&
      request.row >= aggregateRange.rowStart &&
      request.row <= aggregateRange.rowEnd
    ) {
      return true
    }
  }
  return directCriteria.criteriaPairs.some((pair) => {
    if (
      pair.range.sheetName === request.sheetName &&
      pair.range.col === request.col &&
      request.row >= pair.range.rowStart &&
      request.row <= pair.range.rowEnd
    ) {
      return true
    }
    if (pair.criterion.kind === 'literal') {
      return false
    }
    return request.inputCellIndex !== undefined && pair.criterion.cellIndex === request.inputCellIndex
  })
}

export function composeSingleDisjointExplicitEventChanges(explicitCellIndex: number, recalculated: U32): U32 {
  if (recalculated.length === 0) {
    return Uint32Array.of(explicitCellIndex)
  }
  const changed = new Uint32Array(recalculated.length + 1)
  changed[0] = explicitCellIndex
  changed.set(recalculated, 1)
  return changed
}

export function hasCompleteDirectFormulaDeltas(postRecalcDirectFormulaIndices: DirectFormulaIndexCollection): boolean {
  return postRecalcDirectFormulaIndices.hasCompleteDeltas()
}

export function directFormulaChangesAreDisjointFromInputs(
  changedInputArray: U32,
  changedInputCount: number,
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
): boolean {
  for (let index = 0; index < changedInputCount; index += 1) {
    if (postRecalcDirectFormulaIndices.has(changedInputArray[index]!)) {
      return false
    }
  }
  return true
}

export function countDirectFormulaDeltaSkip(
  formulas: { get(cellIndex: number): DirectFormulaRecalcRecord | undefined },
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
  counters: EngineCounters,
): void {
  let sawAggregate = false
  let sawScalar = false
  postRecalcDirectFormulaIndices.forEach((cellIndex) => {
    const formula = formulas.get(cellIndex)
    sawAggregate ||= formula?.directAggregate !== undefined || formula?.directCriteria !== undefined
    sawScalar ||= formula?.directScalar !== undefined
  })
  if (sawAggregate) {
    addEngineCounter(counters, 'directAggregateDeltaOnlyRecalcSkips')
  }
  if (sawScalar) {
    addEngineCounter(counters, 'directScalarDeltaOnlyRecalcSkips')
  }
}

export function canEvaluatePostRecalcDirectFormulasWithoutKernel(
  formulas: { get(cellIndex: number): DirectFormulaRecalcRecord | undefined },
  postRecalcDirectFormulaIndices: DirectFormulaIndexCollection,
): boolean {
  if (postRecalcDirectFormulaIndices.size === 0) {
    return false
  }
  let canEvaluate = true
  postRecalcDirectFormulaIndices.forEach((cellIndex) => {
    const formula = formulas.get(cellIndex)
    if (
      formula?.directAggregate === undefined &&
      formula?.directCriteria === undefined &&
      formula?.directLookup === undefined &&
      formula?.directScalar === undefined
    ) {
      canEvaluate = false
    }
  })
  return canEvaluate
}

export function lookupImpactCacheKey(sheetId: number, col: number): string {
  return `${sheetId}:${col}`
}

export function aggregateColumnDependencyKey(sheetId: number, col: number): number {
  return sheetId * MAX_COLS + col
}
