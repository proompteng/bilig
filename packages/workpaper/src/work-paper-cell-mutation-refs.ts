import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import { canSkipDimensionUpdateAfterLiteralMutation, countPotentialNewTrackedCellMutations } from './work-paper-tracked-event-helpers.js'

export interface WorkPaperCellMutationApplyOptions {
  readonly captureUndo?: boolean
  readonly potentialNewCells?: number
  readonly source?: 'local' | 'restore'
  readonly returnUndoOps?: boolean
  readonly reuseRefs?: boolean
  readonly skipDimensionUpdate?: boolean
}

export interface WorkPaperCellMutationApplyRuntime {
  readonly isEvaluationSuspended: () => boolean
  readonly appendSuspendedCellMutationRefs: (refs: readonly EngineCellMutationRef[]) => void
  readonly addSuspendedCellMutationPotentialNewCells: (amount: number) => void
  readonly applyCellMutationsAtWithOptions: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly updateSheetDimensionsAfterCellMutationRefs: (refs: readonly EngineCellMutationRef[]) => void
}

export function applyWorkPaperCellMutationRefs(
  runtime: WorkPaperCellMutationApplyRuntime,
  refs: readonly EngineCellMutationRef[],
  options: WorkPaperCellMutationApplyOptions,
): void {
  if (runtime.isEvaluationSuspended() && (options.source ?? 'local') === 'local') {
    runtime.appendSuspendedCellMutationRefs(cloneWorkPaperCellMutationRefs(refs))
    runtime.addSuspendedCellMutationPotentialNewCells(options.potentialNewCells ?? countPotentialNewTrackedCellMutations(refs))
    return
  }
  runtime.applyCellMutationsAtWithOptions(refs, engineApplyOptions(options))
  if (options.skipDimensionUpdate !== true && !canSkipDimensionUpdateAfterLiteralMutation(refs, options.potentialNewCells)) {
    runtime.updateSheetDimensionsAfterCellMutationRefs(refs)
  }
}

export function applyQueuedWorkPaperCellMutationRefs(args: {
  readonly refs: readonly EngineCellMutationRef[]
  readonly potentialNewCells: number
  readonly applyCellMutationsAtWithOptions: (refs: readonly EngineCellMutationRef[], options: WorkPaperCellMutationApplyOptions) => void
  readonly updateSheetDimensionsAfterCellMutationRefs: (refs: readonly EngineCellMutationRef[]) => void
  readonly alwaysUpdateDimensions?: boolean
}): void {
  args.applyCellMutationsAtWithOptions(args.refs, {
    captureUndo: true,
    potentialNewCells: args.potentialNewCells,
    source: 'local',
    returnUndoOps: false,
    reuseRefs: true,
  })
  if (args.alwaysUpdateDimensions || !canSkipDimensionUpdateAfterLiteralMutation(args.refs, args.potentialNewCells)) {
    args.updateSheetDimensionsAfterCellMutationRefs(args.refs)
  }
}

function cloneWorkPaperCellMutationRefs(refs: readonly EngineCellMutationRef[]): EngineCellMutationRef[] {
  return refs.map((ref) => ({
    sheetId: ref.sheetId,
    ...(ref.cellIndex !== undefined ? { cellIndex: ref.cellIndex } : {}),
    mutation:
      ref.mutation.kind === 'setCellValue'
        ? {
            kind: 'setCellValue',
            row: ref.mutation.row,
            col: ref.mutation.col,
            value: ref.mutation.value,
          }
        : ref.mutation.kind === 'setCellFormula'
          ? {
              kind: 'setCellFormula',
              row: ref.mutation.row,
              col: ref.mutation.col,
              formula: ref.mutation.formula,
            }
          : {
              kind: 'clearCell',
              row: ref.mutation.row,
              col: ref.mutation.col,
            },
  }))
}

function engineApplyOptions(options: WorkPaperCellMutationApplyOptions): WorkPaperCellMutationApplyOptions {
  if (options.skipDimensionUpdate === undefined) {
    return options
  }
  const { skipDimensionUpdate: _skipDimensionUpdate, ...engineOptions } = options
  return engineOptions
}
