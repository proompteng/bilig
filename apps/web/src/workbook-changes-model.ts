import {
  deriveWorkbookActorHistoryState,
  normalizeWorkbookChangeRowModel,
  workbookChangeRowHistoryRangeSource,
  workbookHistoryRangesOverlap,
  type WorkbookChangeUndoBundle,
  type WorkbookChangeRange,
} from '@bilig/zero-sync'
import { formatWorkbookCollaboratorLabel } from './workbook-presence-model.js'

export interface WorkbookChangeRow {
  readonly revision: number
  readonly actorUserId: string
  readonly clientMutationId: string | null
  readonly eventKind: string
  readonly summary: string
  readonly sheetId: number | null
  readonly sheetName: string | null
  readonly anchorAddress: string | null
  readonly rangeJson: WorkbookChangeRange | null
  readonly rangeJsonInvalid: boolean
  readonly undoBundleJson: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
  readonly createdAt: number
}

export interface WorkbookChangeEntry {
  readonly revision: number
  readonly actorUserId: string
  readonly actorLabel: string
  readonly clientMutationId: string | null
  readonly eventKind: string
  readonly summary: string
  readonly sheetName: string | null
  readonly address: string | null
  readonly targetLabel: string | null
  readonly createdAt: number
  readonly isJumpable: boolean
  readonly canRevert: boolean
  readonly canRedo: boolean
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
}

export interface WorkbookHistoryState {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly undoRevision: number | null
  readonly redoRevision: number | null
}

function normalizeWorkbookChangeRow(value: unknown): WorkbookChangeRow | null {
  const model = normalizeWorkbookChangeRowModel(value)
  if (!model) {
    return null
  }
  return {
    revision: model.revision,
    actorUserId: model.actorUserId,
    clientMutationId: model.clientMutationId,
    eventKind: model.eventKind,
    summary: model.summary,
    sheetId: model.sheetId,
    sheetName: model.sheetName,
    anchorAddress: model.anchorAddress,
    rangeJson: model.rangeJson,
    rangeJsonInvalid: model.rangeJsonInvalid,
    undoBundleJson: model.undoBundleJson,
    revertedByRevision: model.revertedByRevision,
    revertsRevision: model.revertsRevision,
    createdAt: model.createdAt,
  }
}

export function normalizeWorkbookChangeRows(value: unknown): readonly WorkbookChangeRow[] {
  if (!Array.isArray(value)) {
    return []
  }
  return value.flatMap((entry) => {
    const row = normalizeWorkbookChangeRow(entry)
    return row ? [row] : []
  })
}

function formatChangeTarget(range: WorkbookChangeRange | null, fallbackAddress: string | null): string | null {
  if (range) {
    return range.startAddress === range.endAddress
      ? `${range.sheetName}!${range.startAddress}`
      : `${range.sheetName}!${range.startAddress}:${range.endAddress}`
  }
  return fallbackAddress ? fallbackAddress : null
}

export function formatWorkbookChangeDay(createdAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
  }).format(new Date(createdAt))
}

export function formatWorkbookChangeTime(createdAt: number): string {
  return new Intl.DateTimeFormat(undefined, {
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(createdAt))
}

export function selectWorkbookChangeEntries(input: {
  readonly rows: readonly WorkbookChangeRow[]
  readonly knownSheetNames: readonly string[]
}): readonly WorkbookChangeEntry[] {
  const knownSheetNames = new Set(input.knownSheetNames)
  return input.rows.map((row) => {
    const targetLabel = formatChangeTarget(row.rangeJson, row.anchorAddress)
    const sheetName = row.sheetName ?? row.rangeJson?.sheetName ?? null
    const address = row.anchorAddress ?? row.rangeJson?.startAddress ?? null
    const hasLaterOverlap = input.rows.some(
      (candidate) =>
        candidate.revision > row.revision &&
        candidate.revertedByRevision === null &&
        workbookHistoryRangesOverlap(workbookChangeRowHistoryRangeSource(row), workbookChangeRowHistoryRangeSource(candidate)),
    )
    return {
      revision: row.revision,
      actorUserId: row.actorUserId,
      actorLabel: formatWorkbookCollaboratorLabel(row.actorUserId),
      clientMutationId: row.clientMutationId,
      eventKind: row.eventKind,
      summary: row.summary,
      sheetName,
      address,
      targetLabel,
      createdAt: row.createdAt,
      isJumpable: typeof sheetName === 'string' && typeof address === 'string' && knownSheetNames.has(sheetName),
      canRevert: row.undoBundleJson !== null && row.revertedByRevision === null && row.eventKind !== 'revertChange' && !hasLaterOverlap,
      canRedo: row.undoBundleJson !== null && row.revertedByRevision === null && row.eventKind === 'revertChange',
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
    } satisfies WorkbookChangeEntry
  })
}

export function selectWorkbookHistoryState(input: {
  readonly rows: readonly WorkbookChangeRow[]
  readonly currentUserId: string
}): WorkbookHistoryState {
  const history = deriveWorkbookActorHistoryState({
    actorUserId: input.currentUserId,
    rows: input.rows.map((row) => ({
      revision: row.revision,
      actorUserId: row.actorUserId,
      eventKind: row.eventKind,
      undoBundleJson: row.undoBundleJson,
      revertedByRevision: row.revertedByRevision,
      revertsRevision: row.revertsRevision,
      ...workbookChangeRowHistoryRangeSource(row),
    })),
  })

  return {
    canUndo: history.canUndo,
    canRedo: history.canRedo,
    undoRevision: history.undoRevision,
    redoRevision: history.redoRevision,
  }
}
