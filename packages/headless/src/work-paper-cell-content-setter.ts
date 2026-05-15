import type { EngineCellMutationRef, EngineExistingNumericCellMutationResult, SheetRecord, SpreadsheetEngine } from '@bilig/core'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { WorkPaperOperationError } from './work-paper-errors.js'
import { assertRowAndColumn, isFormulaContent, isWorkPaperSheetMatrix } from './work-paper-runtime-helpers.js'
import { buildWorkPaperRawCellMutation } from './work-paper-literal-mutation-queue.js'
import type { WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperChange, WorkPaperConfig, WorkPaperSheet } from './work-paper-types.js'

interface ExistingNumericMutationEngine {
  readonly tryApplyExistingNumericCellMutationAt?: (request: {
    readonly sheetId: number
    readonly row: number
    readonly col: number
    readonly cellIndex: number
    readonly value: number
  }) => EngineExistingNumericCellMutationResult | null
}

export interface WorkPaperSetCellContentsRuntime {
  readonly assertNotDisposed: () => void
  readonly getConfig: () => Pick<WorkPaperConfig, 'maxRows' | 'maxColumns'>
  readonly getEngine: () => SpreadsheetEngine
  readonly sheetRecord: (sheetId: number) => SheetRecord
  readonly getVisibleCellIndexInSheet: (sheet: SheetRecord, row: number, col: number) => number | undefined
  readonly isEvaluationSuspended: () => boolean
  readonly getBatchDepth: () => number
  readonly enqueueSuspendedLiteralMutation: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly enqueueDeferredBatchLiteral: (
    sheetId: number,
    row: number,
    col: number,
    content: RawCellContent,
    cellIndex: number | undefined,
  ) => boolean
  readonly trySetExistingNumericCellContentsWithTrackedFastPath: (args: {
    readonly sheet: SheetRecord
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly value: number
  }) => WorkPaperChange[] | null
  readonly trySetExistingLiteralCellContentsWithTrackedFastPath: (args: {
    readonly sheet: SheetRecord
    readonly address: WorkPaperCellAddress
    readonly cellIndex: number
    readonly value: Exclude<RawCellContent, null>
  }) => WorkPaperChange[] | null
  readonly flushPendingBatchOps: () => void
  readonly rewriteFormulaForStorage: (formula: string, ownerSheetId: number) => string
  readonly applyCellMutationRefs: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly canUseTrackedMutationFastPath: () => boolean
  readonly captureTrackedChangesWithoutVisibilityCache: (
    mutate: () => void,
    options: {
      readonly singleLiteralChange?: {
        readonly address: WorkPaperCellAddress
        readonly cellIndex?: number
        readonly isPhysicalSheet: boolean
        readonly sheetName: string
        readonly value: RawCellContent
      }
    },
  ) => WorkPaperChange[]
  readonly captureChanges: (mutate: () => void) => WorkPaperChange[]
  readonly isItPossibleToSetCellContents: (address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet) => boolean
  readonly applyMatrixContents: (address: WorkPaperCellAddress, content: WorkPaperSheet) => void
}

export function setWorkPaperCellContents(
  runtime: WorkPaperSetCellContentsRuntime,
  address: WorkPaperCellAddress,
  content: RawCellContent | WorkPaperSheet,
): WorkPaperChange[] {
  runtime.assertNotDisposed()
  const sheet = runtime.sheetRecord(address.sheet)
  assertRowAndColumn(address.row, 'address.row')
  assertRowAndColumn(address.col, 'address.col')
  if (!isWorkPaperSheetMatrix(content)) {
    const config = runtime.getConfig()
    if (address.row >= (config.maxRows ?? MAX_ROWS) || address.col >= (config.maxColumns ?? MAX_COLS)) {
      throw new WorkPaperOperationError('Cell contents cannot be set')
    }
    const visibleCellIndex = runtime.getVisibleCellIndexInSheet(sheet, address.row, address.col)
    if (
      runtime.isEvaluationSuspended() &&
      runtime.enqueueSuspendedLiteralMutation(address.sheet, address.row, address.col, content, visibleCellIndex)
    ) {
      return []
    }
    if (
      runtime.getBatchDepth() !== 0 &&
      runtime.enqueueDeferredBatchLiteral(address.sheet, address.row, address.col, content, visibleCellIndex)
    ) {
      return []
    }
    if (typeof content === 'number' && visibleCellIndex !== undefined) {
      const fastPathChanges = runtime.trySetExistingNumericCellContentsWithTrackedFastPath({
        sheet,
        address,
        cellIndex: visibleCellIndex,
        value: content,
      })
      if (fastPathChanges !== null) {
        return fastPathChanges
      }
    }
    if (content !== null && !isFormulaContent(content) && typeof content !== 'number' && visibleCellIndex !== undefined) {
      const fastPathChanges = runtime.trySetExistingLiteralCellContentsWithTrackedFastPath({
        sheet,
        address,
        cellIndex: visibleCellIndex,
        value: content,
      })
      if (fastPathChanges !== null) {
        return fastPathChanges
      }
    }
    const mutate = () => {
      runtime.flushPendingBatchOps()
      const existingNumericMutationEngine = runtime.getEngine() as ExistingNumericMutationEngine
      if (
        typeof content === 'number' &&
        visibleCellIndex !== undefined &&
        sheet.structureVersion === 1 &&
        existingNumericMutationEngine.tryApplyExistingNumericCellMutationAt?.({
          sheetId: address.sheet,
          row: address.row,
          col: address.col,
          cellIndex: visibleCellIndex,
          value: content,
        })
      ) {
        return
      }
      const mutation = buildWorkPaperRawCellMutation({
        row: address.row,
        col: address.col,
        content,
        rewriteFormulaForStorage: (formula) => runtime.rewriteFormulaForStorage(formula, address.sheet),
      })
      runtime.applyCellMutationRefs(
        [{ sheetId: address.sheet, mutation, ...(visibleCellIndex !== undefined ? { cellIndex: visibleCellIndex } : {}) }],
        {
          captureUndo: true,
          potentialNewCells: content === null || visibleCellIndex !== undefined ? 0 : 1,
          source: 'local',
          returnUndoOps: false,
          reuseRefs: true,
        },
      )
    }
    if (runtime.canUseTrackedMutationFastPath()) {
      const captureOptions: {
        singleLiteralChange?: {
          address: WorkPaperCellAddress
          cellIndex?: number
          isPhysicalSheet: boolean
          sheetName: string
          value: RawCellContent
        }
      } = {}
      if (!isFormulaContent(content)) {
        captureOptions.singleLiteralChange = {
          address: { sheet: address.sheet, row: address.row, col: address.col },
          ...(visibleCellIndex === undefined ? {} : { cellIndex: visibleCellIndex }),
          isPhysicalSheet: sheet.structureVersion === 1,
          sheetName: sheet.name,
          value: content,
        }
      }
      return runtime.captureTrackedChangesWithoutVisibilityCache(mutate, captureOptions)
    }
    return runtime.captureChanges(mutate)
  }
  if (!runtime.isItPossibleToSetCellContents(address, content)) {
    throw new WorkPaperOperationError('Cell contents cannot be set')
  }
  return runtime.captureChanges(() => {
    runtime.flushPendingBatchOps()
    runtime.applyMatrixContents(address, content)
  })
}
