import { FormulaMode, ValueTag, type CellValue, type LiteralInput, type RecalcMetrics } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import { makeCellEntity } from '../../entity-ids.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { SheetRecord } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { cellValuesEqual } from './formula-evaluation-helpers.js'
import { makeCompactExistingNumericMutationResult, makeExistingNumericMutationResult } from './operation-change-helpers.js'
import { emitOperationTrackedCellsBatch } from './operation-tracked-event-helpers.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'

const NUMBER_VALUE_TAG = ValueTag.Number
const WASM_FAST_PATH_FORMULA_MODE = FormulaMode.WasmFastPath

export interface OperationTrustedFormulaLeafExistingNumericMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly sheet: SheetRecord
  readonly col: number
  readonly value: number
  readonly hasTrackedEventListeners: boolean
}

export interface OperationFormulaLeafExistingLiteralMutationRequest {
  readonly existingIndex: number
  readonly formulaCellIndex: number
  readonly value: LiteralInput
  readonly hasTrackedEventListeners: boolean
}

export interface OperationFormulaLeafExistingNumericFastPathArgs {
  readonly state: Pick<
    CreateEngineOperationServiceArgs['state'],
    'workbook' | 'strings' | 'wasm' | 'formulas' | 'counters' | 'events' | 'setLastMetrics'
  >
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly writeTrustedExistingNumericLiteralToCell: (existingIndex: number, sheet: SheetRecord, col: number, value: number) => void
  readonly evaluateFormulaCell: (formulaCellIndex: number) => readonly number[]
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export interface OperationFormulaLeafExistingLiteralFastPathArgs {
  readonly state: OperationFormulaLeafExistingNumericFastPathArgs['state']
  readonly getSingleEntityDependent: (entityId: number) => number
  readonly writeFastPathLiteralToExistingCell: (existingIndex: number, value: LiteralInput) => void
  readonly evaluateFormulaCell: (formulaCellIndex: number) => readonly number[]
  readonly deferSingleCellKernelSync: (cellIndex: number) => void
  readonly makeSingleLiteralSkipMetrics: () => RecalcMetrics
}

export function tryApplyTrustedFormulaLeafExistingNumericMutation(
  args: OperationFormulaLeafExistingNumericFastPathArgs,
  request: OperationTrustedFormulaLeafExistingNumericMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formula = getApplicableFormulaLeaf(args, request.formulaCellIndex)
  if (!formula) {
    return null
  }

  const cellStore = args.state.workbook.cellStore
  const beforeFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  args.writeTrustedExistingNumericLiteralToCell(request.existingIndex, request.sheet, request.col, request.value)
  const evaluatedByWasm = evaluateFormulaLeafAfterInputWrite(args, request.existingIndex, request.formulaCellIndex, formula)
  const afterFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  const formulaChanged = !cellValuesEqual(beforeFormulaValue, afterFormulaValue)
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = {
    ...args.makeSingleLiteralSkipMetrics(),
    wasmFormulaCount: evaluatedByWasm ? 1 : 0,
    jsFormulaCount: evaluatedByWasm ? 0 : 1,
  }
  args.state.setLastMetrics(lastMetrics)

  const result = formulaChanged
    ? makeCompactExistingNumericMutationResult(
        request.existingIndex,
        request.formulaCellIndex,
        1,
        afterFormulaValue.tag === NUMBER_VALUE_TAG ? afterFormulaValue.value : undefined,
        {
          row: cellStore.rows[request.formulaCellIndex] ?? 0,
          col: cellStore.cols[request.formulaCellIndex] ?? 0,
        },
      )
    : makeCompactExistingNumericMutationResult(request.existingIndex, undefined, 1)
  if (request.hasTrackedEventListeners) {
    const changedCellIndices = formulaChanged
      ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
      : Uint32Array.of(request.existingIndex)
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return result
}

export function tryApplyFormulaLeafExistingLiteralMutation(
  args: OperationFormulaLeafExistingLiteralFastPathArgs,
  request: OperationFormulaLeafExistingLiteralMutationRequest,
): EngineExistingNumericCellMutationResult | null {
  const formula = getApplicableFormulaLeaf(args, request.formulaCellIndex)
  if (!formula) {
    return null
  }

  const beforeFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  args.writeFastPathLiteralToExistingCell(request.existingIndex, request.value)
  const evaluatedByWasm = evaluateFormulaLeafAfterInputWrite(args, request.existingIndex, request.formulaCellIndex, formula)
  const afterFormulaValue = readFormulaCellValue(args, request.formulaCellIndex)
  const formulaChanged = !cellValuesEqual(beforeFormulaValue, afterFormulaValue)
  addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
  args.deferSingleCellKernelSync(request.existingIndex)
  const lastMetrics = {
    ...args.makeSingleLiteralSkipMetrics(),
    wasmFormulaCount: evaluatedByWasm ? 1 : 0,
    jsFormulaCount: evaluatedByWasm ? 0 : 1,
  }
  args.state.setLastMetrics(lastMetrics)
  const changedCellIndices = formulaChanged
    ? Uint32Array.of(request.existingIndex, request.formulaCellIndex)
    : Uint32Array.of(request.existingIndex)
  if (request.hasTrackedEventListeners) {
    emitOperationTrackedCellsBatch({
      events: args.state.events,
      changedCellIndices,
      metrics: lastMetrics,
    })
  }
  return makeExistingNumericMutationResult(changedCellIndices, 1)
}

function getApplicableFormulaLeaf(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state' | 'getSingleEntityDependent'>,
  formulaCellIndex: number,
): RuntimeFormula | undefined {
  if (formulaCellIndex < 0 || args.getSingleEntityDependent(makeCellEntity(formulaCellIndex)) !== -1) {
    return undefined
  }
  const formula = args.state.formulas.get(formulaCellIndex)
  if (
    !formula ||
    formula.directLookup !== undefined ||
    formula.directAggregate !== undefined ||
    formula.directCriteria !== undefined ||
    formula.directScalar !== undefined ||
    formula.compiled.volatile ||
    formula.compiled.producesSpill ||
    ((args.state.workbook.cellStore.flags[formulaCellIndex] ?? 0) & CellFlags.InCycle) !== 0
  ) {
    return undefined
  }
  return formula
}

function evaluateFormulaLeafAfterInputWrite(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state' | 'evaluateFormulaCell'>,
  existingIndex: number,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): boolean {
  const evaluatedByWasm = tryEvaluateLeafFormulaWithWasm(args, existingIndex, formulaCellIndex, formula)
  if (!evaluatedByWasm) {
    args.evaluateFormulaCell(formulaCellIndex)
  }
  return evaluatedByWasm
}

function tryEvaluateLeafFormulaWithWasm(
  args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>,
  existingIndex: number,
  formulaCellIndex: number,
  formula: RuntimeFormula,
): boolean {
  if (formula.compiled.mode !== WASM_FAST_PATH_FORMULA_MODE || !args.state.wasm.ready) {
    return false
  }
  args.state.wasm.syncStringPool(args.state.strings.exportLayout())
  args.state.wasm.syncFromStore(
    args.state.workbook.cellStore,
    formula.dependencyIndices.length > 0 ? formula.dependencyIndices : Uint32Array.of(existingIndex),
  )
  const formulaIndices = Uint32Array.of(formulaCellIndex)
  args.state.wasm.evalBatch(formulaIndices)
  args.state.wasm.syncToStore(args.state.workbook.cellStore, formulaIndices, args.state.strings, (changedCellIndex: number) =>
    args.state.workbook.notifyCellValueWritten(changedCellIndex),
  )
  return true
}

function readFormulaCellValue(args: Pick<OperationFormulaLeafExistingNumericFastPathArgs, 'state'>, formulaCellIndex: number): CellValue {
  return args.state.workbook.cellStore.getValue(formulaCellIndex, (stringId) => (stringId === 0 ? '' : args.state.strings.get(stringId)))
}
