import {
  deriveWorkbookActorHistoryState,
  workbookChangeRowHistoryRangeSource,
  type WorkbookChangeRange,
  type WorkbookChangeUndoBundle,
  type WorkbookEventKind,
} from '@bilig/zero-sync'

interface WorkbookActorHistoryRecord {
  readonly revision: number
  readonly actorUserId: string
  readonly eventKind: WorkbookEventKind
  readonly undoBundle: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
  readonly sheetName: string | null
  readonly anchorAddress: string | null
  readonly range: WorkbookChangeRange | null
  readonly rangeInvalid: boolean
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
      ...workbookChangeRowHistoryRangeSource({
        sheetName: row.sheetName,
        anchorAddress: row.anchorAddress,
        rangeJson: row.range,
        rangeJsonInvalid: row.rangeInvalid,
      }),
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
      ...workbookChangeRowHistoryRangeSource({
        sheetName: row.sheetName,
        anchorAddress: row.anchorAddress,
        rangeJson: row.range,
        rangeJsonInvalid: row.rangeInvalid,
      }),
    })),
  }).redoRevision
}
