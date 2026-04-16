import { Effect } from 'effect'
import type { EngineOp } from '@bilig/workbook-domain'
import type { EdgeArena, EdgeSlice } from '../../edge-arena.js'
import { definedNameValuesEqual, renameDefinedNameValueSheet } from '../../engine-metadata-utils.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { createInitialRecalcMetrics, createInitialSelectionState } from '../runtime-state.js'
import { EngineMaintenanceError } from '../errors.js'

function maintenanceErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function estimatePotentialNewCells(ops: readonly EngineOp[]): number {
  let count = 0
  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index]!
    if (op.kind === 'setCellValue' || op.kind === 'setCellFormula' || op.kind === 'setCellFormat') {
      count += 1
    }
  }
  return count
}

export interface EngineMaintenanceService {
  readonly captureSheetCellState: (sheetName: string) => Effect.Effect<EngineOp[], EngineMaintenanceError>
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineMaintenanceError>
  readonly captureColumnRangeCellState: (
    sheetName: string,
    start: number,
    count: number,
  ) => Effect.Effect<EngineOp[], EngineMaintenanceError>
  readonly rewriteDefinedNamesForSheetRename: (oldSheetName: string, newSheetName: string) => Effect.Effect<void, EngineMaintenanceError>
  readonly estimatePotentialNewCells: (ops: readonly EngineOp[]) => Effect.Effect<number, EngineMaintenanceError>
  readonly resetWorkbook: (workbookName?: string) => Effect.Effect<void, EngineMaintenanceError>
}

export function createEngineMaintenanceService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'workbook'
    | 'formulas'
    | 'ranges'
    | 'entityVersions'
    | 'sheetDeleteVersions'
    | 'undoStack'
    | 'redoStack'
    | 'setSelection'
    | 'setSyncState'
    | 'getLastMetrics'
    | 'setLastMetrics'
  >
  readonly edgeArena: EdgeArena
  readonly reverseState: {
    reverseCellEdges: Array<EdgeSlice | undefined>
    reverseRangeEdges: Array<EdgeSlice | undefined>
    reverseDefinedNameEdges: Map<string, Set<number>>
    reverseTableEdges: Map<string, Set<number>>
    reverseSpillEdges: Map<string, Set<number>>
    reverseAggregateColumnEdges: Map<number, Set<number>>
    reverseExactLookupColumnEdges: Map<number, EdgeSlice>
    reverseSortedLookupColumnEdges: Map<number, EdgeSlice>
  }
  readonly pivotOutputOwners: Map<number, string>
  readonly captureSheetCellState: (sheetName: string) => EngineOp[]
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly setWasmProgramSyncPending: (next: boolean) => void
  readonly setMaterializedCellCount: (next: number) => void
  readonly scheduleWasmProgramSync: () => void
  readonly resetWasmState: () => void
}): EngineMaintenanceService {
  const rewriteDefinedNamesForSheetRenameNow = (oldSheetName: string, newSheetName: string): void => {
    args.state.workbook.listDefinedNames().forEach((record) => {
      const nextValue = renameDefinedNameValueSheet(record.value, oldSheetName, newSheetName)
      if (!definedNameValuesEqual(record.value, nextValue)) {
        args.state.workbook.setDefinedName(record.name, nextValue)
      }
    })
  }

  const resetWorkbookNow = (workbookName = 'Workbook'): void => {
    const previousBatchId = args.state.getLastMetrics().batchId
    args.state.workbook.reset(workbookName)
    args.state.formulas.clear()
    args.reverseState.reverseCellEdges.length = 0
    args.reverseState.reverseRangeEdges.length = 0
    args.reverseState.reverseDefinedNameEdges.clear()
    args.reverseState.reverseTableEdges.clear()
    args.reverseState.reverseSpillEdges.clear()
    args.reverseState.reverseAggregateColumnEdges.clear()
    args.reverseState.reverseExactLookupColumnEdges.clear()
    args.reverseState.reverseSortedLookupColumnEdges.clear()
    args.pivotOutputOwners.clear()
    args.state.ranges.reset()
    args.edgeArena.reset()
    args.state.entityVersions.clear()
    args.state.sheetDeleteVersions.clear()
    args.state.undoStack.length = 0
    args.state.redoStack.length = 0
    args.state.setSelection(createInitialSelectionState())
    args.state.setSyncState('local-only')
    args.state.setLastMetrics({
      ...createInitialRecalcMetrics(),
      batchId: previousBatchId,
    })
    args.setWasmProgramSyncPending(false)
    args.setMaterializedCellCount(0)
    args.resetWasmState()
    args.scheduleWasmProgramSync()
  }

  return {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => args.captureSheetCellState(sheetName),
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to capture sheet cell state', cause),
            cause,
          }),
      })
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => args.captureRowRangeCellState(sheetName, start, count),
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to capture row range cell state', cause),
            cause,
          }),
      })
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => args.captureColumnRangeCellState(sheetName, start, count),
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to capture column range cell state', cause),
            cause,
          }),
      })
    },
    rewriteDefinedNamesForSheetRename(oldSheetName, newSheetName) {
      return Effect.try({
        try: () => {
          rewriteDefinedNamesForSheetRenameNow(oldSheetName, newSheetName)
        },
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to rewrite defined names for sheet rename', cause),
            cause,
          }),
      })
    },
    estimatePotentialNewCells(ops) {
      return Effect.try({
        try: () => estimatePotentialNewCells(ops),
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to estimate potential new cells', cause),
            cause,
          }),
      })
    },
    resetWorkbook(workbookName) {
      return Effect.try({
        try: () => {
          resetWorkbookNow(workbookName)
        },
        catch: (cause) =>
          new EngineMaintenanceError({
            message: maintenanceErrorMessage('Failed to reset workbook', cause),
            cause,
          }),
      })
    },
  }
}
