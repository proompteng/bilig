import {
  CellFlags,
  type EngineExistingNumericCellMutationResult,
  type SheetRecord,
  type SpreadsheetEngine,
} from '@bilig/core/headless-runtime'
import { MAX_COLS, MAX_ROWS, ValueTag, type CellValue, type LiteralInput } from '@bilig/protocol'
import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import { readCompactSecondChangedValue, trackedEventFromExistingNumericMutationResult } from './work-paper-tracked-event-helpers.js'
import { isFormulaContent, isWorkPaperSheetMatrix, scalarValueFromLiteral } from './work-paper-runtime-helpers.js'
import type { TrackedEngineEvent } from './tracked-engine-event-refs.js'
import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange, WorkPaperConfig, WorkPaperSheet } from './work-paper-types.js'

const FAST_DIRECT_EXISTING_LITERAL_FLAGS =
  CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput

type ExistingLiteralMutationEngine = SpreadsheetEngine & {
  readonly tryApplyExistingNumericCellMutationAt?: (request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly value: number
    readonly emitTracked?: boolean
    readonly trustedExistingNumericLiteral?: boolean
    readonly oldNumericValue?: number
  }) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyExistingLiteralCellMutationAt?: (request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly value: LiteralInput
    readonly emitTracked?: boolean
  }) => EngineExistingNumericCellMutationResult | null
}

export interface WorkPaperExistingLiteralDirectFastPathRuntime {
  readonly computeTrackedChangesWithoutVisibilityCache: (
    events: readonly TrackedEngineEvent[],
    options: { readonly preferLazyPublicChanges?: boolean },
  ) => WorkPaperChange[]
  readonly getBatchDepth: () => number
  readonly getConfig: () => WorkPaperConfig
  readonly getEngine: () => SpreadsheetEngine
  readonly hasNamedExpressions: () => boolean
  readonly hasPendingBatchOps: () => boolean
  readonly hasPendingLazyTrackedChanges: () => boolean
  readonly hasTrackedEngineEvents: () => boolean
  readonly hasValuesUpdatedListeners: () => boolean
  readonly isDisposed: () => boolean
  readonly isEvaluationSuspended: () => boolean
  readonly clearTrackedEngineEvents: () => void
  readonly messageOf: (error: unknown, fallback: string) => string
  readonly readSingleTrackedCellChange: (cellIndex: number) => WorkPaperCellChange | undefined
  readonly trackedA1: (row: number, col: number) => string
}

export function trySetExistingLiteralWorkPaperCellContentsDirectFastPath(
  runtime: WorkPaperExistingLiteralDirectFastPathRuntime,
  address: WorkPaperCellAddress,
  content: LiteralInput | WorkPaperSheet,
): WorkPaperChange[] | null {
  const config = runtime.getConfig()
  if (
    runtime.isDisposed() ||
    isWorkPaperSheetMatrix(content) ||
    content === null ||
    isFormulaContent(content) ||
    !Number.isInteger(address.row) ||
    !Number.isInteger(address.col) ||
    address.row < 0 ||
    address.col < 0 ||
    address.row >= (config.maxRows ?? MAX_ROWS) ||
    address.col >= (config.maxColumns ?? MAX_COLS) ||
    runtime.isEvaluationSuspended() ||
    runtime.getBatchDepth() !== 0 ||
    runtime.hasNamedExpressions() ||
    runtime.hasValuesUpdatedListeners()
  ) {
    return null
  }
  const engine = runtime.getEngine()
  const sheet = engine.workbook.getSheetById(address.sheet)
  if (!sheet || sheet.structureVersion !== 1) {
    return null
  }
  const cellIndex = sheet.grid.getPhysical(address.row, address.col)
  if (cellIndex === -1) {
    return null
  }
  const cellStore = engine.workbook.cellStore
  if (
    cellStore.sheetIds[cellIndex] !== address.sheet ||
    cellStore.rows[cellIndex] !== address.row ||
    cellStore.cols[cellIndex] !== address.col ||
    (cellStore.formulaIds[cellIndex] ?? 0) !== 0 ||
    ((cellStore.flags[cellIndex] ?? 0) & FAST_DIRECT_EXISTING_LITERAL_FLAGS) !== 0
  ) {
    return null
  }
  if (typeof content === 'number' && cellStore.tags[cellIndex] !== ValueTag.Number) {
    return null
  }
  const mutationEngine = engine as ExistingLiteralMutationEngine
  if (
    (typeof content === 'number' && typeof mutationEngine.tryApplyExistingNumericCellMutationAt !== 'function') ||
    (typeof content !== 'number' && typeof mutationEngine.tryApplyExistingLiteralCellMutationAt !== 'function')
  ) {
    return null
  }

  if (runtime.hasPendingLazyTrackedChanges() || runtime.hasTrackedEngineEvents() || runtime.hasPendingBatchOps()) {
    return null
  }
  let result: EngineExistingNumericCellMutationResult | null = null
  try {
    if (typeof content === 'number') {
      result =
        mutationEngine.tryApplyExistingNumericCellMutationAt?.({
          sheetId: address.sheet,
          row: address.row,
          col: address.col,
          cellIndex,
          value: content,
          emitTracked: false,
          trustedExistingNumericLiteral: true,
          oldNumericValue: cellStore.numbers[cellIndex] ?? 0,
        }) ?? null
    } else {
      result =
        mutationEngine.tryApplyExistingLiteralCellMutationAt?.({
          sheetId: address.sheet,
          row: address.row,
          col: address.col,
          cellIndex,
          value: content,
          emitTracked: false,
        }) ?? null
    }
  } catch (error) {
    if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
  }
  if (!result) {
    return null
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.clearTrackedEngineEvents()
  }
  return materializeExistingLiteralDirectChanges(runtime, result, {
    address,
    cellIndex,
    content,
    sheet,
  })
}

function materializeExistingLiteralDirectChanges(
  runtime: WorkPaperExistingLiteralDirectFastPathRuntime,
  result: EngineExistingNumericCellMutationResult,
  request: {
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly content: LiteralInput
    readonly sheet: SheetRecord
  },
): WorkPaperChange[] {
  return (
    tryReadCompactExistingLiteralDirectChanges(runtime, result, request) ??
    runtime.computeTrackedChangesWithoutVisibilityCache([trackedEventFromExistingNumericMutationResult(result)], {
      preferLazyPublicChanges: true,
    })
  )
}

function tryReadCompactExistingLiteralDirectChanges(
  runtime: WorkPaperExistingLiteralDirectFastPathRuntime,
  result: EngineExistingNumericCellMutationResult,
  request: {
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly content: LiteralInput
    readonly sheet: SheetRecord
  },
): WorkPaperChange[] | null {
  const changedCellIndices = result.changedCellIndices
  if (changedCellIndices !== undefined) {
    if (changedCellIndices.length > 4) {
      return null
    }
    const changes: WorkPaperChange[] = []
    for (let index = 0; index < changedCellIndices.length; index += 1) {
      const change = runtime.readSingleTrackedCellChange(changedCellIndices[index]!)
      if (change === undefined) {
        return null
      }
      changes.push(change)
    }
    return changes
  }
  const changedCellCount = result.changedCellCount ?? 0
  if (
    changedCellCount === 0 ||
    changedCellCount > 2 ||
    result.firstChangedCellIndex === undefined ||
    result.firstChangedCellIndex !== request.cellIndex
  ) {
    return null
  }
  const firstChange = {
    kind: 'cell' as const,
    address: { sheet: request.address.sheet, row: request.address.row, col: request.address.col },
    sheetName: request.sheet.name,
    a1: runtime.trackedA1(request.address.row, request.address.col),
    newValue: scalarValueFromLiteral(request.content),
  }
  if (changedCellCount === 1) {
    return [firstChange]
  }
  if (result.secondChangedCellIndex === undefined) {
    return null
  }
  const cellStore = runtime.getEngine().workbook.cellStore
  const secondRow = result.secondChangedRow ?? cellStore.rows[result.secondChangedCellIndex]
  const secondCol = result.secondChangedCol ?? cellStore.cols[result.secondChangedCellIndex]
  if (secondRow === undefined || secondCol === undefined) {
    return null
  }
  const secondChangedValue = readCompactSecondChangedValue(result)
  const secondValue: CellValue =
    secondChangedValue ??
    (result.secondChangedNumericValue === undefined
      ? cellStore.getValue(result.secondChangedCellIndex, (stringId) => (stringId === 0 ? '' : runtime.getEngine().strings.get(stringId)))
      : { tag: ValueTag.Number, value: result.secondChangedNumericValue })
  const secondChange = {
    kind: 'cell' as const,
    address: { sheet: request.address.sheet, row: secondRow, col: secondCol },
    sheetName: request.sheet.name,
    a1: runtime.trackedA1(secondRow, secondCol),
    newValue: secondValue,
  }
  return [firstChange, secondChange]
}
