import { ValueTag, type CellValue } from '@bilig/protocol'

export interface ArrayValue {
  kind: 'array'
  values: CellValue[]
  rows: number
  cols: number
}

export interface RangeLikeValue {
  kind: 'range'
  values: CellValue[]
  rows: number
  cols: number
  refKind: 'cells' | 'rows' | 'cols'
}

export type EvaluationResult = CellValue | ArrayValue

export function isArrayValue(value: EvaluationResult): value is ArrayValue {
  return 'kind' in value && value.kind === 'array'
}

export function scalarFromEvaluationResult(value: EvaluationResult): CellValue {
  if (!isArrayValue(value)) {
    return value
  }
  return value.values[0] ?? { tag: ValueTag.Empty }
}
