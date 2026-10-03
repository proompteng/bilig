import { workPaperHistoryTopIsCellMutations, type WorkPaperHistoryRecord } from './work-paper-history.js'
import type { WorkPaperChange } from './work-paper-types.js'

export function applyWorkPaperHistoryOperation(args: {
  readonly getStack: () => readonly WorkPaperHistoryRecord[]
  readonly canUseTrackedMutationFastPath: () => boolean
  readonly captureTrackedChangesWithoutVisibilityCache: (
    mutate: () => void,
    options: { readonly preservePendingTrackedPositions?: boolean },
  ) => WorkPaperChange[]
  readonly captureChanges: (mutate: () => void, options: { readonly preservePendingTrackedPositions?: boolean }) => WorkPaperChange[]
  readonly applyOperation: () => boolean
  readonly createMissingOperationError: () => Error
  readonly invalidateAllSheetDimensions: () => void
}): WorkPaperChange[] {
  const preservesPositions = !workPaperHistoryTopIsCellMutations(args.getStack())
  const mutate = () => {
    if (!args.applyOperation()) {
      throw args.createMissingOperationError()
    }
    args.invalidateAllSheetDimensions()
  }
  return args.canUseTrackedMutationFastPath()
    ? args.captureTrackedChangesWithoutVisibilityCache(mutate, { preservePendingTrackedPositions: preservesPositions })
    : args.captureChanges(mutate, { preservePendingTrackedPositions: preservesPositions })
}
