import { ErrorCode, ValueTag, type EngineChangedCell } from '@bilig/protocol'
import type { EngineOpBatch } from '@bilig/workbook-domain'
import { formatAddress, type CompiledFormula, type FormulaNode, type ParsedCellReferenceInfo } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import type { EngineCellMutationRef } from '../../cell-mutations-at.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import { batchOpOrder, markBatchApplied, type OpOrder } from '../../replica-state.js'
import type { EngineRuntimeState, U32 } from '../runtime-state.js'
import type { DirectScalarCurrentOperand } from './direct-formula-index-collection.js'
import { unwrapDirectScalarBinaryNode } from './formula-binding-direct-scalar.js'
import { rethrowFatalFormulaBindingError } from './formula-binding-error-policy.js'
import { emitCellMutationFastPathBatchResult } from './operation-fast-path-batch-result.js'
import { tagTrustedPhysicalTrackedChanges } from './operation-change-helpers.js'
import { freshMatrixOverlapsFormulaDependencies } from './operation-fresh-matrix-dependency-overlap.js'
import {
  attachFreshDenseDirectAggregateMatrixCells,
  materializeFreshMatrixAxisIds,
} from './operation-fresh-direct-aggregate-matrix-helpers.js'
import {
  createFreshFormulaInstanceList,
  registerFreshBoundFormulaFamilyRun,
} from './operation-fresh-direct-aggregate-formula-batch-records.js'
import type { CreateEngineOperationServiceArgs } from './operation-service-types.js'
import type { FreshDirectScalarFormulaBindingMember } from './formula-binding-service-types.js'

const EMPTY_CHANGED_CELLS = new Uint32Array(0)
const FRESH_DIRECT_SCALAR_FORMULA_BATCH_MIN_SIZE = 32

type FastPathState = Pick<
  EngineRuntimeState,
  | 'workbook'
  | 'wasm'
  | 'ranges'
  | 'strings'
  | 'events'
  | 'formulas'
  | 'counters'
  | 'replicaState'
  | 'getLastMetrics'
  | 'setLastMetrics'
  | 'getSyncClientConnection'
  | 'trackReplicaVersions'
>

interface FreshDirectScalarFormulaEntry extends FreshDirectScalarFormulaBindingMember {
  readonly result: DirectScalarCurrentOperand
}

interface FreshDirectScalarMatrixBatch {
  readonly sheet: NonNullable<ReturnType<FastPathState['workbook']['getSheetById']>>
  readonly sheetId: number
  readonly rowStart: number
  readonly rowCount: number
  readonly colStart: number
  readonly inputColCount: number
  readonly formulaCol: number
  readonly values: Float64Array
  readonly formulaEntries: readonly FreshDirectScalarFormulaEntry[]
}

export interface OperationFreshDirectScalarFormulaBatchFastPathArgs {
  readonly state: FastPathState
  readonly emitBatch: (batch: EngineOpBatch) => void
  readonly setCellEntityVersion: (sheetName: string, address: string, order: OpOrder) => void
  readonly hasTrackedExactLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedSortedLookupDependents: (sheetId: number, col: number) => boolean
  readonly hasTrackedDirectRangeDependents: (sheetId: number, col: number) => boolean
  readonly hasRegionFormulaSubscriptionsOverlappingRange?: (
    sheetId: number,
    rowStart: number,
    rowEnd: number,
    colStart: number,
    colEnd: number,
  ) => boolean
  readonly getRegionFormulaSubscriptionCount?: () => number
  readonly hasRegionFormulaSubscriptionsIntersectingRect:
    | ((sheetId: number, rowStart: number, rowEnd: number, colStart: number, colEnd: number) => boolean)
    | undefined
  readonly bindFreshDirectScalarFormulaRun: NonNullable<CreateEngineOperationServiceArgs['bindFreshDirectScalarFormulaRun']>
  readonly registerFreshFormulaFamilyRun: CreateEngineOperationServiceArgs['registerFreshFormulaFamilyRun']
  readonly upsertFormulaFamilyRun: CreateEngineOperationServiceArgs['upsertFormulaFamilyRun']
  readonly upsertFreshFormulaInstances: CreateEngineOperationServiceArgs['upsertFreshFormulaInstances']
  readonly compileTemplateFormula: NonNullable<CreateEngineOperationServiceArgs['compileTemplateFormula']>
  readonly materializeDeferredStructuralFormulaSources: () => void
  readonly checkEvaluationBudget: (stepCost?: number) => void
  readonly beginMutationCollection: () => void
  readonly ensureRecalcScratchCapacity: (size: number) => void
  readonly resetMaterializedCellScratch: (expectedSize: number) => void
  readonly getBatchMutationDepth: () => number
  readonly setBatchMutationDepth: (next: number) => void
  readonly markInputChanged: (cellIndex: number, count: number) => number
  readonly markExplicitChanged: (cellIndex: number, count: number) => number
  readonly getChangedInputBuffer: () => U32
  readonly deferKernelSync: (cellIndices: readonly number[] | U32) => void
  readonly captureChangedCells: (changedCellIndices: readonly number[] | U32) => readonly EngineChangedCell[]
  readonly applyDirectFormulaCurrentResult: (cellIndex: number, result: DirectScalarCurrentOperand) => boolean
}

export function createOperationFreshDirectScalarFormulaBatchFastPath(args: OperationFreshDirectScalarFormulaBatchFastPathArgs): {
  readonly tryApplyFreshDirectScalarFormulaMatrixBatch: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => boolean
} {
  const tryApplyFreshDirectScalarFormulaMatrixBatch = (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ): boolean => {
    const firstRef = refs[0]
    if (
      source !== 'local' ||
      firstRef === undefined ||
      refs.length < FRESH_DIRECT_SCALAR_FORMULA_BATCH_MIN_SIZE ||
      potentialNewCells !== refs.length ||
      args.state.workbook.metadata.definedNames.size > 0
    ) {
      return false
    }
    if (args.state.trackReplicaVersions && batch === null) {
      return false
    }
    args.checkEvaluationBudget()
    const matrix = collectFreshDirectScalarMatrixBatch(args, refs, firstRef)
    if (matrix === null || freshMatrixOverlapsFormulaDependencies(args, matrix)) {
      return false
    }

    args.materializeDeferredStructuralFormulaSources()
    args.beginMutationCollection()
    args.ensureRecalcScratchCapacity(args.state.workbook.cellStore.size + refs.length + 1)
    args.resetMaterializedCellScratch(refs.length)

    const hasGeneralEventListeners = args.state.events.hasListeners()
    const hasTrackedEventListeners = args.state.events.hasTrackedListeners()
    const hasWatchedCellListeners = args.state.events.hasCellListeners()
    const requiresChangedSet = hasGeneralEventListeners || hasTrackedEventListeners || hasWatchedCellListeners
    const ensureRowId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'row')
    const ensureColId = args.state.workbook.createLogicalAxisIdEnsurer(matrix.sheetId, 'column')
    const totalColCount = matrix.inputColCount + 1
    const rowIds = materializeFreshMatrixAxisIds(matrix.rowCount, matrix.rowStart, ensureRowId)
    const columnIds = materializeFreshMatrixAxisIds(totalColCount, matrix.colStart, ensureColId)
    const firstCellIndex = args.state.workbook.cellStore.allocateDenseRowMajorAtReserved(
      matrix.sheetId,
      matrix.rowStart,
      matrix.rowCount,
      matrix.colStart,
      totalColCount,
    )
    attachFreshDenseDirectAggregateMatrixCells(matrix.sheet, firstCellIndex, matrix.rowStart, matrix.colStart, rowIds, columnIds)
    const formulaCellIndices = new Uint32Array(matrix.rowCount)
    const formulaInstances = args.upsertFreshFormulaInstances === undefined ? undefined : createFreshFormulaInstanceList(matrix.rowCount)
    let changedInputCount = 0
    let explicitChangedCount = 0
    args.setBatchMutationDepth(args.getBatchMutationDepth() + 1)
    try {
      args.state.workbook.withBatchedColumnVersionUpdates(() => {
        let valueIndex = 0
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          args.checkEvaluationBudget(matrix.inputColCount + 1)
          const row = matrix.rowStart + rowOffset
          for (let colOffset = 0; colOffset < matrix.inputColCount; colOffset += 1) {
            const col = matrix.colStart + colOffset
            const cellIndex = firstCellIndex + rowOffset * totalColCount + colOffset
            writeFreshNumericLiteralToCellStore(args, cellIndex, matrix.values[valueIndex]!)
            changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
            if (requiresChangedSet) {
              explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
            }
            if (args.state.trackReplicaVersions && batch !== null) {
              args.setCellEntityVersion(matrix.sheet.name, formatAddress(row, col), batchOpOrder(batch, valueIndex))
            }
            valueIndex += 1
          }
        }
        for (let rowOffset = 0; rowOffset < matrix.rowCount; rowOffset += 1) {
          args.checkEvaluationBudget()
          const entry = matrix.formulaEntries[rowOffset]!
          const cellIndex = firstCellIndex + rowOffset * totalColCount + matrix.inputColCount
          formulaCellIndices[rowOffset] = cellIndex
          if (formulaInstances !== undefined) {
            formulaInstances[rowOffset] = {
              cellIndex,
              sheetName: matrix.sheet.name,
              row: entry.row,
              col: entry.col,
              source: entry.source,
              templateId: entry.templateId,
            }
          }
          args.applyDirectFormulaCurrentResult(cellIndex, entry.result)
          changedInputCount = args.markInputChanged(cellIndex, changedInputCount)
          if (requiresChangedSet) {
            explicitChangedCount = args.markExplicitChanged(cellIndex, explicitChangedCount)
          }
          if (args.state.trackReplicaVersions && batch !== null) {
            args.setCellEntityVersion(
              matrix.sheet.name,
              formatAddress(entry.row, entry.col),
              batchOpOrder(batch, matrix.values.length + rowOffset),
            )
          }
        }
        args.bindFreshDirectScalarFormulaRun({
          sheetId: matrix.sheetId,
          ownerSheetName: matrix.sheet.name,
          cellIndices: formulaCellIndices,
          members: matrix.formulaEntries,
        })
        if (args.upsertFormulaFamilyRun !== undefined) {
          registerFreshBoundFormulaFamilyRun(args, matrix.sheetId, matrix.formulaEntries, formulaCellIndices)
        }
        if (formulaInstances !== undefined) {
          args.upsertFreshFormulaInstances?.(formulaInstances)
        }
        const writtenColumns = new Uint32Array(matrix.inputColCount)
        for (let index = 0; index < matrix.inputColCount; index += 1) {
          writtenColumns[index] = matrix.colStart + index
        }
        args.state.workbook.notifyColumnsWritten(matrix.sheetId, writtenColumns)
      })
    } finally {
      args.setBatchMutationDepth(args.getBatchMutationDepth() - 1)
    }
    if (batch) {
      markBatchApplied(args.state.replicaState, batch)
    }
    const changedInputArray = args.getChangedInputBuffer().subarray(0, changedInputCount)
    args.deferKernelSync(changedInputArray)
    if (requiresChangedSet && hasTrackedEventListeners && matrix.values.length > 0 && matrix.values.length < changedInputArray.length) {
      tagTrustedPhysicalTrackedChanges(changedInputArray, matrix.sheetId, matrix.values.length)
    }
    addEngineCounter(args.state.counters, 'kernelSyncOnlyRecalcSkips')
    addEngineCounter(args.state.counters, 'directFormulaKernelSyncOnlyRecalcSkips')
    emitCellMutationFastPathBatchResult({
      state: args.state,
      changed: requiresChangedSet ? changedInputArray : EMPTY_CHANGED_CELLS,
      changedInputCount,
      explicitChangedCount,
      hasGeneralEventListeners,
      hasTrackedEventListeners,
      hasWatchedCellListeners,
      captureChangedCells: args.captureChangedCells,
      batch,
      emitBatch: args.emitBatch,
    })
    return true
  }

  return { tryApplyFreshDirectScalarFormulaMatrixBatch }
}

function collectFreshDirectScalarMatrixBatch(
  args: OperationFreshDirectScalarFormulaBatchFastPathArgs,
  refs: readonly EngineCellMutationRef[],
  firstRef: EngineCellMutationRef,
): FreshDirectScalarMatrixBatch | null {
  const firstMutation = firstRef.mutation
  if (firstMutation.kind !== 'setCellValue' || typeof firstMutation.value !== 'number' || Object.is(firstMutation.value, -0)) {
    return null
  }
  const sheet = args.state.workbook.getSheetById(firstRef.sheetId)
  if (!sheet) {
    return null
  }

  let firstFormulaRefIndex = -1
  for (let index = 0; index < refs.length; index += 1) {
    args.checkEvaluationBudget()
    if (refs[index]!.mutation.kind === 'setCellFormula') {
      firstFormulaRefIndex = index
      break
    }
  }
  if (firstFormulaRefIndex <= 0) {
    return null
  }

  const values = new Float64Array(firstFormulaRefIndex)
  let currentRow = firstMutation.row
  let currentWidth = 0
  let inputColCount = 0
  let rowCount = 1
  for (let refIndex = 0; refIndex < firstFormulaRefIndex; refIndex += 1) {
    args.checkEvaluationBudget()
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellValue' ||
      typeof mutation.value !== 'number' ||
      Object.is(mutation.value, -0)
    ) {
      return null
    }
    if (mutation.row === currentRow) {
      if (mutation.col !== firstMutation.col + currentWidth) {
        return null
      }
      currentWidth += 1
    } else {
      if (mutation.row !== currentRow + 1 || mutation.col !== firstMutation.col || currentWidth === 0) {
        return null
      }
      if (inputColCount === 0) {
        inputColCount = currentWidth
      } else if (currentWidth !== inputColCount) {
        return null
      }
      currentRow = mutation.row
      currentWidth = 1
      rowCount += 1
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined
    ) {
      return null
    }
    values[refIndex] = mutation.value
  }
  if (inputColCount === 0) {
    inputColCount = currentWidth
  } else if (currentWidth !== inputColCount) {
    return null
  }
  if (rowCount < 2 || inputColCount < 1 || rowCount * inputColCount !== firstFormulaRefIndex) {
    return null
  }
  if (refs.length - firstFormulaRefIndex !== rowCount) {
    return null
  }
  const formulaCol = firstMutation.col + inputColCount
  for (let col = firstMutation.col; col <= formulaCol; col += 1) {
    if (
      args.hasTrackedExactLookupDependents(firstRef.sheetId, col) ||
      args.hasTrackedSortedLookupDependents(firstRef.sheetId, col) ||
      args.hasTrackedDirectRangeDependents(firstRef.sheetId, col)
    ) {
      return null
    }
  }

  const formulaEntries: FreshDirectScalarFormulaEntry[] = []
  for (let refIndex = firstFormulaRefIndex; refIndex < refs.length; refIndex += 1) {
    args.checkEvaluationBudget()
    const ref = refs[refIndex]!
    const mutation = ref.mutation
    const rowOffset = refIndex - firstFormulaRefIndex
    if (
      ref.sheetId !== firstRef.sheetId ||
      ref.cellIndex !== undefined ||
      mutation.kind !== 'setCellFormula' ||
      mutation.row !== firstMutation.row + rowOffset ||
      mutation.col !== formulaCol
    ) {
      return null
    }
    if (
      sheet.grid.getPhysical(mutation.row, mutation.col) !== -1 ||
      sheet.logical.getVisibleCell(mutation.row, mutation.col) !== undefined
    ) {
      return null
    }
    let template: ReturnType<OperationFreshDirectScalarFormulaBatchFastPathArgs['compileTemplateFormula']>
    try {
      template = args.compileTemplateFormula(mutation.formula, mutation.row, mutation.col)
    } catch (error) {
      rethrowFatalFormulaBindingError(error)
      return null
    }
    const compiled = template.compiled
    if (!canUseFreshDirectScalarMatrixFormula(compiled)) {
      return null
    }
    const result = evaluateFreshDirectScalarMatrixRow({
      compiled,
      inputColCount,
      matrixColStart: firstMutation.col,
      ownerSheetName: sheet.name,
      row: mutation.row,
      rowOffset,
      values,
    })
    if (result === undefined) {
      return null
    }
    formulaEntries.push({
      row: mutation.row,
      col: mutation.col,
      source: mutation.formula,
      compiled,
      templateId: template.templateId,
      result,
    })
  }

  return {
    sheet,
    sheetId: firstRef.sheetId,
    rowStart: firstMutation.row,
    rowCount,
    colStart: firstMutation.col,
    inputColCount,
    formulaCol,
    values,
    formulaEntries,
  }
}

function canUseFreshDirectScalarMatrixFormula(compiled: CompiledFormula): boolean {
  if (
    compiled.volatile ||
    compiled.producesSpill ||
    compiled.symbolicNames.length !== 0 ||
    compiled.symbolicTables.length !== 0 ||
    compiled.symbolicSpills.length !== 0 ||
    compiled.symbolicRanges.length !== 0 ||
    compiled.directAggregateCandidate !== undefined
  ) {
    return false
  }
  const node = unwrapDirectScalarBinaryNode(compiled.optimizedAst).node
  return (
    (node.kind === 'BinaryExpr' &&
      (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/') &&
      (node.left.kind === 'CellRef' || node.left.kind === 'NumberLiteral') &&
      (node.right.kind === 'CellRef' || node.right.kind === 'NumberLiteral')) ||
    (node.kind === 'CallExpr' && node.callee.trim().toUpperCase() === 'ABS' && node.args.length === 1 && node.args[0]?.kind === 'CellRef')
  )
}

function evaluateFreshDirectScalarMatrixRow(input: {
  readonly compiled: CompiledFormula
  readonly ownerSheetName: string
  readonly row: number
  readonly rowOffset: number
  readonly matrixColStart: number
  readonly inputColCount: number
  readonly values: Float64Array
}): DirectScalarCurrentOperand | undefined {
  const refs = input.compiled.parsedSymbolicRefs
  if (refs === undefined || refs.length !== input.compiled.symbolicRefs.length) {
    return undefined
  }
  let refIndex = 0
  const readOperand = (node: FormulaNode): DirectScalarCurrentOperand | undefined => {
    if (node.kind === 'NumberLiteral') {
      return Number.isFinite(node.value) ? { kind: 'number', value: node.value } : undefined
    }
    if (node.kind !== 'CellRef') {
      return undefined
    }
    const ref = refs[refIndex]
    refIndex += 1
    return ref === undefined ? undefined : readFreshDirectScalarMatrixCell(input, ref)
  }

  const unwrapped = unwrapDirectScalarBinaryNode(input.compiled.optimizedAst)
  const node = unwrapped.node
  let result: DirectScalarCurrentOperand | undefined
  if (node.kind === 'CallExpr' && node.callee.trim().toUpperCase() === 'ABS' && node.args.length === 1) {
    const operand = readOperand(node.args[0]!)
    result = operand?.kind === 'number' ? { kind: 'number', value: Math.abs(operand.value) } : operand
  } else if (
    node.kind === 'BinaryExpr' &&
    (node.operator === '+' || node.operator === '-' || node.operator === '*' || node.operator === '/')
  ) {
    const left = readOperand(node.left)
    const right = readOperand(node.right)
    if (left === undefined || right === undefined) {
      return undefined
    }
    if (left.kind === 'error') {
      result = left
    } else if (right.kind === 'error') {
      result = right
    } else {
      let value: number
      switch (node.operator) {
        case '+':
          value = left.value + right.value
          break
        case '-':
          value = left.value - right.value
          break
        case '*':
          value = left.value * right.value
          break
        case '/':
          if (right.value === 0) {
            result = { kind: 'error', code: ErrorCode.Div0 }
          } else {
            value = left.value / right.value
            result = { kind: 'number', value }
          }
          break
      }
      if (result === undefined) {
        result = { kind: 'number', value: value! }
      }
    }
  } else {
    return undefined
  }
  if (refIndex !== refs.length || result === undefined) {
    return undefined
  }
  return result.kind === 'number' && unwrapped.resultOffset !== undefined
    ? { kind: 'number', value: result.value + unwrapped.resultOffset }
    : result
}

function readFreshDirectScalarMatrixCell(
  input: {
    readonly ownerSheetName: string
    readonly row: number
    readonly rowOffset: number
    readonly matrixColStart: number
    readonly inputColCount: number
    readonly values: Float64Array
  },
  ref: ParsedCellReferenceInfo,
): DirectScalarCurrentOperand | undefined {
  if (
    (ref.sheetName !== undefined && ref.sheetName !== input.ownerSheetName) ||
    ref.row === undefined ||
    ref.col === undefined ||
    ref.row !== input.row ||
    ref.col < input.matrixColStart ||
    ref.col >= input.matrixColStart + input.inputColCount
  ) {
    return undefined
  }
  return {
    kind: 'number',
    value: input.values[input.rowOffset * input.inputColCount + ref.col - input.matrixColStart]!,
  }
}

function writeFreshNumericLiteralToCellStore(
  args: OperationFreshDirectScalarFormulaBatchFastPathArgs,
  cellIndex: number,
  value: number,
): void {
  const cellStore = args.state.workbook.cellStore
  cellStore.flags[cellIndex] = (cellStore.flags[cellIndex] ?? 0) & ~(CellFlags.SpillChild | CellFlags.PivotOutput)
  cellStore.versions[cellIndex] = (cellStore.versions[cellIndex] ?? 0) + 1
  cellStore.stringIds[cellIndex] = 0
  cellStore.tags[cellIndex] = ValueTag.Number
  cellStore.numbers[cellIndex] = value
  cellStore.errors[cellIndex] = ErrorCode.None
}
