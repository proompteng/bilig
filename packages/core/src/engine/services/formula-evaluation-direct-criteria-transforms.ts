import { getBuiltin, isArrayValue } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { errorValue } from '../../engine-value-utils.js'
import type { RuntimeDirectCriteriaResultTransform, RuntimeFormula } from '../runtime-state.js'
import type { RuntimeColumnView } from './runtime-column-store-service.js'

type CellValueReader = (cellIndex: number | undefined) => CellValue

const roundBuiltin = getBuiltin('ROUND')

const isEmptyCellGuardValue = (value: CellValue): boolean =>
  value.tag === ValueTag.Empty || (value.tag === ValueTag.String && value.value === '')

const applyDirectCriteriaResultTransform = (
  readCellValueByIndex: CellValueReader,
  transform: RuntimeDirectCriteriaResultTransform,
  current: CellValue,
): CellValue => {
  if (transform.kind === 'if-error') {
    return current.tag === ValueTag.Error ? transform.fallback : current
  }
  if (transform.kind === 'if-empty-cell') {
    const guard = readCellValueByIndex(transform.cellIndex)
    if (guard.tag === ValueTag.Error) {
      return guard
    }
    return isEmptyCellGuardValue(guard) ? transform.fallback : current
  }
  const rounded = roundBuiltin?.(current, transform.digits) ?? errorValue(ErrorCode.Name)
  return isArrayValue(rounded) ? errorValue(ErrorCode.Value) : rounded
}

const applyDirectCriteriaResultTransformsFrom = (
  readCellValueByIndex: CellValueReader,
  formula: RuntimeFormula,
  value: CellValue,
  startIndex: number,
): CellValue => {
  const transforms = formula.directCriteria?.resultTransforms
  if (!transforms || transforms.length === 0) {
    return value
  }
  let current = value
  for (let index = startIndex; index < transforms.length; index += 1) {
    current = applyDirectCriteriaResultTransform(readCellValueByIndex, transforms[index]!, current)
  }
  return current
}

export const applyDirectCriteriaResultTransforms = (
  readCellValueByIndex: CellValueReader,
  formula: RuntimeFormula,
  value: CellValue,
): CellValue => applyDirectCriteriaResultTransformsFrom(readCellValueByIndex, formula, value, 0)

export const tryEvaluateDirectCriteriaTransformShortCircuit = (
  readCellValueByIndex: CellValueReader,
  formula: RuntimeFormula,
): CellValue | undefined => {
  const transforms = formula.directCriteria?.resultTransforms
  if (!transforms || transforms.length === 0) {
    return undefined
  }
  for (let index = 0; index < transforms.length; index += 1) {
    const transform = transforms[index]!
    if (transform.kind !== 'if-empty-cell') {
      continue
    }
    const guard = readCellValueByIndex(transform.cellIndex)
    if (guard.tag === ValueTag.Error) {
      return applyDirectCriteriaResultTransformsFrom(readCellValueByIndex, formula, guard, index + 1)
    }
    if (isEmptyCellGuardValue(guard)) {
      return applyDirectCriteriaResultTransformsFrom(readCellValueByIndex, formula, transform.fallback, index + 1)
    }
  }
  return undefined
}

const readColumnViewTag = (view: RuntimeColumnView, offset: number): ValueTag => view.readTagAt(offset) as ValueTag

export const numericLikeValueInView = (view: RuntimeColumnView, offset: number): number | undefined => {
  switch (readColumnViewTag(view, offset)) {
    case ValueTag.Number:
      return view.readNumberAt(offset)
    case ValueTag.Boolean:
      return view.readNumberAt(offset) !== 0 ? 1 : 0
    case ValueTag.Empty:
      return 0
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return undefined
  }
}

export const strictNumericAggregateCandidateInView = (view: RuntimeColumnView, offset: number): number | undefined =>
  readColumnViewTag(view, offset) === ValueTag.Number ? view.readNumberAt(offset) : undefined
