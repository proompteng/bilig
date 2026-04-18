import { deriveWorkbookActorHistoryState, type WorkbookChangeUndoBundle } from '@bilig/zero-sync'

interface WorkbookActorHistoryRecord {
  readonly revision: number
  readonly actorUserId: string
  readonly eventKind: string
  readonly undoBundle: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
}

export function selectLatestUndoableWorkbookChangeRevision(input: {
  readonly actorUserId: string
  readonly rows: readonly WorkbookActorHistoryRecord[]
}): number | null {
  return deriveWorkbookActorHistoryState({
    actorUserId: input.actorUserId,
    rows: input.rows.map((row) => ({
      revision: row.revision,
      actorUserId: row.actorUserId,
      eventKind: row.eventKind,
      undoBundleJson: row.undoBundle,
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    })),
  }).undoRevision
}

export function selectLatestRedoableWorkbookChangeRevision(input: {
  readonly actorUserId: string
  readonly rows: readonly WorkbookActorHistoryRecord[]
}): number | null {
  return deriveWorkbookActorHistoryState({
    actorUserId: input.actorUserId,
    rows: input.rows.map((row) => ({
      revision: row.revision,
      actorUserId: row.actorUserId,
      eventKind: row.eventKind,
      undoBundleJson: row.undoBundle,
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    })),
  }).redoRevision
}
