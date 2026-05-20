import type { WorkbookChangeUndoBundle, WorkbookEventKind } from './workbook-events.js'
import type { WorkbookChangeRange, WorkbookChangeRangeScope } from './workbook-change-range.js'
import { normalizeWorkbookChangeRange, normalizeWorkbookChangeRangeBounds } from './workbook-change-range.js'

export type WorkbookHistoryRange = WorkbookChangeRange

export interface WorkbookHistoryRangeSource {
  readonly sheetName?: string | null | undefined
  readonly anchorAddress?: string | null | undefined
  readonly rangeJson?: WorkbookHistoryRange | null | undefined
  readonly rangeJsonInvalid?: boolean | null | undefined
}

export interface WorkbookHistoryStateRow {
  readonly revision: number
  readonly actorUserId: string
  readonly eventKind: WorkbookEventKind
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
  readonly scope: WorkbookChangeRangeScope
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
  if (source.rangeJsonInvalid) {
    return null
  }
  const range =
    source.rangeJson === undefined || source.rangeJson === null
      ? source.sheetName && source.anchorAddress
        ? { sheetName: source.sheetName, startAddress: source.anchorAddress, endAddress: source.anchorAddress }
        : source.sheetName
          ? { sheetName: source.sheetName, startAddress: 'A1', endAddress: 'A1', scope: 'sheet' as const }
          : null
      : normalizeWorkbookChangeRange(source.rangeJson)
  if (!range) {
    return null
  }
  const bounds = normalizeWorkbookChangeRangeBounds(range)
  if (!bounds) {
    return null
  }
  return {
    sheetName: bounds.sheetName,
    startRow: bounds.rowStart,
    endRow: bounds.rowEnd,
    startCol: bounds.colStart,
    endCol: bounds.colEnd,
    scope: bounds.scope,
  }
}

function rangesOverlap(left: NormalizedWorkbookHistoryRange | null, right: NormalizedWorkbookHistoryRange | null): boolean {
  if (!left || !right) {
    return true
  }
  if (left.sheetName !== right.sheetName) {
    return false
  }
  if (left.scope === 'sheet' || right.scope === 'sheet') {
    return true
  }
  if (left.scope === 'rows' && right.scope === 'columns') {
    return true
  }
  if (left.scope === 'columns' && right.scope === 'rows') {
    return true
  }
  if (left.scope === 'rows' || right.scope === 'rows') {
    return left.startRow <= right.endRow && left.endRow >= right.startRow
  }
  if (left.scope === 'columns' || right.scope === 'columns') {
    return left.startCol <= right.endCol && left.endCol >= right.startCol
  }
  return left.startRow <= right.endRow && left.endRow >= right.startRow && left.startCol <= right.endCol && left.endCol >= right.startCol
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
    if (row.eventKind === 'revertChange') {
      undoStack = removeRevision(undoStack, row.revertsRevision)
      if (row.undoBundleJson !== null && row.revertedByRevision === null) {
        redoStack = pushUnique(redoStack, stackEntry)
      } else {
        redoStack = removeRevision(redoStack, row.revision)
      }
      continue
    }
    if (row.eventKind === 'redoChange') {
      redoStack = removeRevision(redoStack, row.revertsRevision)
      if (row.undoBundleJson !== null && row.revertedByRevision === null) {
        undoStack = pushUnique(undoStack, stackEntry)
      } else {
        undoStack = removeRevision(undoStack, row.revision)
      }
      continue
    }
    redoStack = []
    if (row.undoBundleJson !== null && row.revertedByRevision === null) {
      undoStack = pushUnique(undoStack, stackEntry)
    } else {
      undoStack = removeRevision(undoStack, row.revision)
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
