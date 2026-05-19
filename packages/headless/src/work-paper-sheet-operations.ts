import type { SheetRecord } from '@bilig/core/headless-runtime'
import { formatAddress } from '@bilig/formula'
import type { WorkPaperNamedExpressionValueSnapshot } from './work-paper-named-expression-helpers.js'
import type { VisibilitySnapshot } from './work-paper-visibility-snapshot.js'
import { WorkPaperInvalidArgumentsError, WorkPaperSheetError, WorkPaperSheetNameAlreadyTakenError } from './work-paper-errors.js'
import type { WorkPaperChange, WorkPaperDetailedEventMap, WorkPaperSheet, WorkPaperSheetDimensions } from './work-paper-types.js'

type SheetOperationEvent =
  | {
      readonly eventName: 'sheetRemoved'
      readonly payload: WorkPaperDetailedEventMap['sheetRemoved']
    }
  | {
      readonly eventName: 'sheetRenamed'
      readonly payload: WorkPaperDetailedEventMap['sheetRenamed']
    }

export interface WorkPaperSheetOperationsRuntime {
  readonly assertNotDisposed: () => void
  readonly materializePendingLazyChanges: () => void
  readonly nextSheetName: () => string
  readonly isItPossibleToAddSheet: (name: string) => boolean
  readonly ensureVisibilityCache: () => VisibilitySnapshot
  readonly ensureNamedExpressionValueCache: () => WorkPaperNamedExpressionValueSnapshot
  readonly drainEngineEvents: () => void
  readonly createSheet: (name: string) => void
  readonly clearSheetRecordsCache: () => void
  readonly requireSheetId: (name: string) => number
  readonly cacheSheetDimensions: (sheetId: number, dimensions: WorkPaperSheetDimensions) => void
  readonly shouldSuppressEvents: () => boolean
  readonly queueSheetAddedEvent: (payload: WorkPaperDetailedEventMap['sheetAdded']) => void
  readonly emitSheetAdded: (payload: WorkPaperDetailedEventMap['sheetAdded']) => void
  readonly computeChangesAfterMutation: (
    beforeVisibility: VisibilitySnapshot,
    beforeNames: WorkPaperNamedExpressionValueSnapshot,
  ) => WorkPaperChange[]
  readonly hasValuesUpdatedListeners: () => boolean
  readonly emitValuesUpdated: (changes: WorkPaperChange[]) => void
  readonly isItPossibleToRemoveSheet: (sheetId: number) => boolean
  readonly sheetName: (sheetId: number) => string
  readonly captureChanges: (event: SheetOperationEvent | undefined, mutate: () => void) => WorkPaperChange[]
  readonly deleteSheet: (name: string) => void
  readonly invalidateSheetDimensions: (sheetId: number) => void
  readonly isItPossibleToClearSheet: (sheetId: number) => boolean
  readonly getSheetDimensions: (sheetId: number) => WorkPaperSheetDimensions
  readonly clearRange: (range: { readonly sheetName: string; readonly startAddress: string; readonly endAddress: string }) => void
  readonly isItPossibleToReplaceSheetContent: (sheetId: number, content: WorkPaperSheet) => boolean
  readonly replaceSheetContentInternal: (
    sheetId: number,
    content: WorkPaperSheet,
    options: { readonly duringInitialization: boolean },
  ) => void
  readonly sheetRecord: (sheetId: number) => SheetRecord
  readonly getSheetByName: (name: string) => SheetRecord | undefined
  readonly tryRenameSheetWithoutVisibilitySnapshots: (oldName: string, newName: string) => WorkPaperChange[] | null
  readonly renameSheet: (oldName: string, newName: string) => void
}

export interface WorkPaperSheetOperations {
  readonly addSheet: (sheetName?: string) => string
  readonly removeSheet: (sheetId: number) => WorkPaperChange[]
  readonly clearSheet: (sheetId: number) => WorkPaperChange[]
  readonly setSheetContent: (sheetId: number, content: WorkPaperSheet) => WorkPaperChange[]
  readonly renameSheet: (sheetId: number, nextName: string) => WorkPaperChange[]
}

export function createWorkPaperSheetOperations(runtime: WorkPaperSheetOperationsRuntime): WorkPaperSheetOperations {
  return {
    addSheet(sheetName) {
      runtime.assertNotDisposed()
      runtime.materializePendingLazyChanges()
      const name = sheetName?.trim() || runtime.nextSheetName()
      if (!runtime.isItPossibleToAddSheet(name)) {
        throw new WorkPaperSheetNameAlreadyTakenError(name)
      }
      const beforeVisibility = runtime.ensureVisibilityCache()
      const beforeNames = runtime.ensureNamedExpressionValueCache()
      runtime.drainEngineEvents()
      runtime.createSheet(name)
      runtime.clearSheetRecordsCache()
      const sheetId = runtime.requireSheetId(name)
      runtime.cacheSheetDimensions(sheetId, { width: 0, height: 0 })
      const payload: WorkPaperDetailedEventMap['sheetAdded'] = { sheetId, sheetName: name }
      if (runtime.shouldSuppressEvents()) {
        runtime.queueSheetAddedEvent(payload)
      } else {
        runtime.emitSheetAdded(payload)
      }
      const changes = runtime.computeChangesAfterMutation(beforeVisibility, beforeNames)
      if (!runtime.shouldSuppressEvents() && changes.length > 0 && runtime.hasValuesUpdatedListeners()) {
        runtime.emitValuesUpdated(changes)
      }
      return name
    },

    removeSheet(sheetId) {
      if (!runtime.isItPossibleToRemoveSheet(sheetId)) {
        throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be removed`)
      }
      const sheetName = runtime.sheetName(sheetId)
      return runtime.captureChanges(
        {
          eventName: 'sheetRemoved',
          payload: {
            sheetId,
            sheetName,
            changes: [],
          },
        },
        () => {
          runtime.deleteSheet(sheetName)
          runtime.clearSheetRecordsCache()
          runtime.invalidateSheetDimensions(sheetId)
        },
      )
    },

    clearSheet(sheetId) {
      if (!runtime.isItPossibleToClearSheet(sheetId)) {
        throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be cleared`)
      }
      return runtime.captureChanges(undefined, () => {
        const dimensions = runtime.getSheetDimensions(sheetId)
        if (dimensions.width === 0 || dimensions.height === 0) {
          return
        }
        runtime.clearRange({
          sheetName: runtime.sheetName(sheetId),
          startAddress: 'A1',
          endAddress: formatAddress(dimensions.height - 1, dimensions.width - 1),
        })
        runtime.cacheSheetDimensions(sheetId, { width: 0, height: 0 })
      })
    },

    setSheetContent(sheetId, content) {
      if (!runtime.isItPossibleToReplaceSheetContent(sheetId, content)) {
        throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be replaced`)
      }
      return runtime.captureChanges(undefined, () => {
        runtime.replaceSheetContentInternal(sheetId, content, { duringInitialization: false })
      })
    },

    renameSheet(sheetId, nextName) {
      const sheet = runtime.sheetRecord(sheetId)
      const newName = nextName.trim()
      if (newName.length === 0) {
        throw new WorkPaperInvalidArgumentsError('Sheet name must be non-empty')
      }
      const existing = runtime.getSheetByName(newName)
      if (existing && existing.id !== sheetId) {
        throw new WorkPaperSheetError(`Sheet '${sheetId}' cannot be renamed to '${nextName}'`)
      }
      const oldName = sheet.name
      const fastPathChanges = runtime.tryRenameSheetWithoutVisibilitySnapshots(oldName, newName)
      if (fastPathChanges) {
        return fastPathChanges
      }
      return runtime.captureChanges(
        {
          eventName: 'sheetRenamed',
          payload: {
            sheetId,
            oldName,
            newName,
          },
        },
        () => {
          runtime.renameSheet(oldName, newName)
          runtime.clearSheetRecordsCache()
        },
      )
    },
  }
}
