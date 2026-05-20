import {
  listWorkbookChangesAfterRevision,
  loadLatestRedoableWorkbookChange,
  loadLatestUndoableWorkbookChange,
  loadWorkbookChange,
  type WorkbookChangeStoreConnection,
  type WorkbookChangeRecord,
} from './workbook-change-store.js'
import type { WorkbookHistoryMutationTarget } from './server-mutator-commit.js'
import { workbookChangeRowHistoryRangeSource, workbookHistoryRangesOverlap } from '@bilig/zero-sync'

export async function resolveRevertWorkbookChangeTarget(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly revision: number
  },
): Promise<WorkbookHistoryMutationTarget> {
  const targetChange = await loadWorkbookChange(db, input.documentId, input.revision)
  if (!targetChange) {
    throw new Error('Workbook change was not found')
  }
  if (!targetChange.undoBundle) {
    throw new Error('Workbook change is not revertible')
  }
  if (targetChange.revertedByRevision !== null) {
    throw new Error(`Workbook change was already reverted in r${targetChange.revertedByRevision}`)
  }
  if (targetChange.eventKind === 'revertChange' || targetChange.revertsRevision !== null) {
    throw new Error('Reverting a revert change is not supported')
  }
  await assertNoActiveOverlappingLaterChange(db, input.documentId, targetChange)
  return toWorkbookHistoryMutationTarget(targetChange)
}

export async function resolveUndoLatestWorkbookChangeTarget(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly actorUserId: string
  },
): Promise<WorkbookHistoryMutationTarget> {
  const targetChange = await loadLatestUndoableWorkbookChange(db, input)
  if (!targetChange?.undoBundle) {
    throw new Error('No undoable workbook change was found')
  }
  return toWorkbookHistoryMutationTarget(targetChange)
}

export async function resolveRedoLatestWorkbookChangeTarget(
  db: WorkbookChangeStoreConnection,
  input: {
    readonly documentId: string
    readonly actorUserId: string
  },
): Promise<WorkbookHistoryMutationTarget> {
  const targetChange = await loadLatestRedoableWorkbookChange(db, input)
  if (!targetChange?.undoBundle) {
    throw new Error('No redoable workbook change was found')
  }
  return toWorkbookHistoryMutationTarget(targetChange)
}

function toWorkbookHistoryMutationTarget(change: WorkbookChangeRecord): WorkbookHistoryMutationTarget {
  if (!change.undoBundle) {
    throw new Error('Workbook change is not revertible')
  }
  return {
    revision: change.revision,
    summary: change.summary,
    sheetName: change.sheetName,
    anchorAddress: change.anchorAddress,
    range: change.range,
    undoBundle: change.undoBundle,
  }
}

async function assertNoActiveOverlappingLaterChange(
  db: WorkbookChangeStoreConnection,
  documentId: string,
  targetChange: WorkbookChangeRecord,
): Promise<void> {
  const laterChanges = await listWorkbookChangesAfterRevision(db, {
    documentId,
    revision: targetChange.revision,
  })
  const conflictingChange = laterChanges.find(
    (change) =>
      change.revertedByRevision === null &&
      workbookHistoryRangesOverlap(
        workbookChangeRowHistoryRangeSource({
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          rangeJson: targetChange.range,
          rangeJsonInvalid: targetChange.rangeInvalid,
        }),
        workbookChangeRowHistoryRangeSource({
          sheetName: change.sheetName,
          anchorAddress: change.anchorAddress,
          rangeJson: change.range,
          rangeJsonInvalid: change.rangeInvalid,
        }),
      ),
  )
  if (conflictingChange) {
    throw new Error(`Workbook change cannot be safely reverted after overlapping r${conflictingChange.revision}`)
  }
}
