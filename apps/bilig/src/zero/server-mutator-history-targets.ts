import {
  loadLatestRedoableWorkbookChange,
  loadLatestUndoableWorkbookChange,
  loadWorkbookChange,
  type WorkbookChangeRecord,
} from './workbook-change-store.js'
import type { WorkbookHistoryMutationTarget } from './server-mutator-commit.js'
import type { Queryable } from './store.js'

export async function resolveRevertWorkbookChangeTarget(
  db: Queryable,
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
  return toWorkbookHistoryMutationTarget(targetChange)
}

export async function resolveUndoLatestWorkbookChangeTarget(
  db: Queryable,
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
  db: Queryable,
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
