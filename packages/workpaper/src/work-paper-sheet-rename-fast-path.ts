import { WORKPAPER_PUBLIC_ERROR_NAMES } from './work-paper-config.js'
import { WorkPaperOperationError } from './work-paper-errors.js'
import type { WorkPaperChange } from './work-paper-types.js'

export interface WorkPaperSheetRenameFastPathRuntime {
  readonly canUseMetadataOnlySheetRenameFastPath: () => boolean
  readonly assertNotDisposed: () => void
  readonly hasPendingLazyTrackedChanges: () => boolean
  readonly materializePendingLazyTrackedChanges: () => void
  readonly isTrackedBatchFastPathActive: () => boolean
  readonly downgradeTrackedBatchFastPath: () => void
  readonly hasTrackedEngineEvents: () => boolean
  readonly drainTrackedEngineEvents: () => void
  readonly clearTrackedEngineEvents: () => void
  readonly clearSheetRecordsCache: () => void
  readonly renameSheetMetadataOnly: (oldName: string, newName: string) => boolean
  readonly renameSheet: (oldName: string, newName: string) => void
  readonly withEngineEventCaptureDisabled: (callback: () => void) => void
  readonly messageOf: (error: unknown, fallback: string) => string
}

export function tryRenameWorkPaperSheetWithoutVisibilitySnapshots(
  runtime: WorkPaperSheetRenameFastPathRuntime,
  oldName: string,
  newName: string,
): WorkPaperChange[] | null {
  if (!runtime.canUseMetadataOnlySheetRenameFastPath()) {
    return null
  }
  runtime.assertNotDisposed()
  if (runtime.hasPendingLazyTrackedChanges()) {
    runtime.materializePendingLazyTrackedChanges()
  }
  if (runtime.isTrackedBatchFastPathActive()) {
    runtime.downgradeTrackedBatchFastPath()
  }
  if (!runtime.canUseMetadataOnlySheetRenameFastPath()) {
    return null
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.drainTrackedEngineEvents()
  }
  try {
    if (runtime.renameSheetMetadataOnly(oldName, newName)) {
      runtime.clearSheetRecordsCache()
      runtime.clearTrackedEngineEvents()
    } else {
      runtime.withEngineEventCaptureDisabled(() => {
        runtime.renameSheet(oldName, newName)
        runtime.clearSheetRecordsCache()
      })
    }
  } catch (error) {
    if (error instanceof Error && WORKPAPER_PUBLIC_ERROR_NAMES.has(error.name)) {
      throw error
    }
    throw new WorkPaperOperationError(runtime.messageOf(error, 'Mutation failed'))
  }
  if (runtime.hasTrackedEngineEvents()) {
    runtime.drainTrackedEngineEvents()
  }
  return []
}
