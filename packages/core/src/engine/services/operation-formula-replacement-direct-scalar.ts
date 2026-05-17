import { ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { makeCellEntity } from '../../entity-ids.js'
import type { EngineRuntimeState, RuntimeDirectScalarDescriptor } from '../runtime-state.js'
import type { DirectFormulaIndexCollection, DirectScalarCurrentOperand } from './direct-formula-index-collection.js'

interface OperationFormulaReplacementDirectScalarArgs {
  readonly state: Pick<EngineRuntimeState, 'formulas' | 'workbook'>
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly evaluateDirectScalarCurrentValue: (directScalar: RuntimeDirectScalarDescriptor) => DirectScalarCurrentOperand | undefined
  readonly tryMarkDirectScalarLinearDeltaClosure: (
    cellIndex: number,
    oldValue: CellValue,
    newValue: CellValue,
    collection: DirectFormulaIndexCollection,
  ) => boolean
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
}

interface OperationFormulaReplacementDirectScalarRequest {
  readonly cellIndex: number
  readonly oldNumber: number | undefined
  readonly changedTopology: boolean
  readonly postRecalcDirectFormulaIndices: DirectFormulaIndexCollection
}

export function tryApplyOperationFormulaReplacementAsDirectScalarDeltaRoot(
  args: OperationFormulaReplacementDirectScalarArgs,
  request: OperationFormulaReplacementDirectScalarRequest,
): boolean {
  if (request.changedTopology || request.oldNumber === undefined) {
    return false
  }
  const formula = args.state.formulas.get(request.cellIndex)
  if (
    !formula ||
    formula.directScalar === undefined ||
    formula.compiled.volatile ||
    formula.compiled.producesSpill ||
    ((args.state.workbook.cellStore.flags[request.cellIndex] ?? 0) & CellFlags.InCycle) !== 0
  ) {
    return false
  }
  const result = args.evaluateDirectScalarCurrentValue(formula.directScalar)
  if (result?.kind !== 'number') {
    return false
  }
  const dependent = args.getSingleEntityDependent(makeCellEntity(request.cellIndex))
  if (
    dependent !== -1 &&
    !args.tryMarkDirectScalarLinearDeltaClosure(
      request.cellIndex,
      { tag: ValueTag.Number, value: request.oldNumber },
      { tag: ValueTag.Number, value: result.value },
      request.postRecalcDirectFormulaIndices,
    )
  ) {
    return false
  }
  if (!args.applyDirectFormulaCurrentResult(request.cellIndex, result)) {
    return false
  }
  return true
}
