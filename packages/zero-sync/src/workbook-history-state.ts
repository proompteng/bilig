import type { WorkbookChangeUndoBundle } from './workbook-events.js'
import { parseCellAddress } from '@bilig/formula'

export interface WorkbookHistoryRange {
  readonly sheetName: string
  readonly startAddress: string
  readonly endAddress: string
}

export interface WorkbookHistoryRangeSource {
  readonly sheetName?: string | null | undefined
  readonly anchorAddress?: string | null | undefined
  readonly rangeJson?: WorkbookHistoryRange | null | undefined
}

export interface WorkbookHistoryStateRow {
  readonly revision: number
  readonly actorUserId: string
  readonly eventKind: string
  readonly undoBundleJson: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
  readonly sheetName?: WorkbookHistoryRangeSource['sheetName']
  readonly anchorAddress?: WorkbookHistoryRangeSource['anchorAddress']
  readonly rangeJson?: WorkbookHistoryRangeSource['rangeJson']
}

export interface WorkbookActorHistoryState {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly undoRevision: number | null
  readonly redoRevision: number | null
  readonly undoStack: readonly number[]
  readonly redoStack: readonly number[]
}

interface HistoryStackEntry {
  readonly revision: number
  readonly range: NormalizedWorkbookHistoryRange | null
}

interface NormalizedWorkbookHistoryRange {
  readonly sheetName: string
  readonly startRow: number
  readonly endRow: number
  readonly startCol: number
  readonly endCol: number
}

function pushUnique(stack: readonly HistoryStackEntry[], entry: HistoryStackEntry): HistoryStackEntry[] {
  const next = stack.filter((candidate) => candidate.revision !== entry.revision)
  next.push(entry)
  return next
}

function removeRevision(stack: readonly HistoryStackEntry[], revision: number | null): HistoryStackEntry[] {
  if (revision === null) {
    return [...stack]
  }
  return stack.filter((entry) => entry.revision !== revision)
}

function removeOverlappingEntries(
  stack: readonly HistoryStackEntry[],
  changedRange: NormalizedWorkbookHistoryRange | null,
): HistoryStackEntry[] {
  return stack.filter((entry) => !rangesOverlap(entry.range, changedRange))
}

function normalizeWorkbookHistoryRange(source: WorkbookHistoryRangeSource): NormalizedWorkbookHistoryRange | null {
  const range =
    source.rangeJson ??
    (source.sheetName && source.anchorAddress
      ? { sheetName: source.sheetName, startAddress: source.anchorAddress, endAddress: source.anchorAddress }
      : null)
  if (!range) {
    return null
  }
  try {
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    return {
      sheetName: range.sheetName,
      startRow: Math.min(start.row, end.row),
      endRow: Math.max(start.row, end.row),
      startCol: Math.min(start.col, end.col),
      endCol: Math.max(start.col, end.col),
    }
  } catch {
    return null
  }
}

function rangesOverlap(left: NormalizedWorkbookHistoryRange | null, right: NormalizedWorkbookHistoryRange | null): boolean {
  if (!left || !right) {
    return true
  }
  return (
    left.sheetName === right.sheetName &&
    left.startRow <= right.endRow &&
    left.endRow >= right.startRow &&
    left.startCol <= right.endCol &&
    left.endCol >= right.startCol
  )
}

export function workbookHistoryRangesOverlap(left: WorkbookHistoryRangeSource, right: WorkbookHistoryRangeSource): boolean {
  return rangesOverlap(normalizeWorkbookHistoryRange(left), normalizeWorkbookHistoryRange(right))
}

export function deriveWorkbookActorHistoryState(input: {
  readonly actorUserId: string
  readonly rows: readonly WorkbookHistoryStateRow[]
}): WorkbookActorHistoryState {
  const rows = [...input.rows].toSorted((left, right) => left.revision - right.revision)

  let undoStack: HistoryStackEntry[] = []
  let redoStack: HistoryStackEntry[] = []

  for (const row of rows) {
    const range = normalizeWorkbookHistoryRange(row)
    if (row.actorUserId !== input.actorUserId) {
      if (row.revertedByRevision === null) {
        undoStack = removeOverlappingEntries(undoStack, range)
        redoStack = removeOverlappingEntries(redoStack, range)
      }
      continue
    }
    const stackEntry = { revision: row.revision, range }
    switch (row.eventKind) {
      case 'revertChange': {
        undoStack = removeRevision(undoStack, row.revertsRevision)
        if (row.undoBundleJson !== null && row.revertedByRevision === null) {
          redoStack = pushUnique(redoStack, stackEntry)
        } else {
          redoStack = removeRevision(redoStack, row.revision)
        }
        break
      }
      case 'redoChange': {
        redoStack = removeRevision(redoStack, row.revertsRevision)
        if (row.undoBundleJson !== null && row.revertedByRevision === null) {
          undoStack = pushUnique(undoStack, stackEntry)
        } else {
          undoStack = removeRevision(undoStack, row.revision)
        }
        break
      }
      default: {
        redoStack = []
        if (row.undoBundleJson !== null && row.revertedByRevision === null) {
          undoStack = pushUnique(undoStack, stackEntry)
        } else {
          undoStack = removeRevision(undoStack, row.revision)
        }
        break
      }
    }
  }

  return {
    canUndo: undoStack.length > 0,
    canRedo: redoStack.length > 0,
    undoRevision: undoStack.at(-1)?.revision ?? null,
    redoRevision: redoStack.at(-1)?.revision ?? null,
    undoStack: undoStack.map((entry) => entry.revision),
    redoStack: redoStack.map((entry) => entry.revision),
  }
}
