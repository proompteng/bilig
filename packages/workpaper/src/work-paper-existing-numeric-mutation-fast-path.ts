import type { EngineExistingNumericCellMutationsRef, SpreadsheetEngine } from '@bilig/core/headless-runtime'
import type { WorkPaperCellMutationApplyOptions } from './work-paper-cell-mutation-refs.js'

export function tryApplyExistingNumericCellMutationsAtWithOptions(
  engine: SpreadsheetEngine,
  record: EngineExistingNumericCellMutationsRef,
  options: WorkPaperCellMutationApplyOptions,
): boolean {
  if (
    options.captureUndo !== true ||
    options.source !== 'local' ||
    options.returnUndoOps !== false ||
    (options.potentialNewCells ?? 0) !== 0
  ) {
    return false
  }
  return engine.tryApplyExistingNumericCellMutationsAt(record)
}
