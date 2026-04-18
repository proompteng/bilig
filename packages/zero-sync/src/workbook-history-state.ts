import type { WorkbookChangeUndoBundle } from './workbook-events.js'

export interface WorkbookHistoryStateRow {
  readonly revision: number
  readonly actorUserId: string
  readonly eventKind: string
  readonly undoBundleJson: WorkbookChangeUndoBundle | null
  readonly revertedByRevision: number | null
  readonly revertsRevision: number | null
}

export interface WorkbookActorHistoryState {
  readonly canUndo: boolean
  readonly canRedo: boolean
  readonly undoRevision: number | null
  readonly redoRevision: number | null
  readonly undoStack: readonly number[]
  readonly redoStack: readonly number[]
}

function pushUnique(stack: number[], revision: number): number[] {
  const next = stack.filter((entry) => entry !== revision)
  next.push(revision)
  return next
}

function removeRevision(stack: number[], revision: number | null): number[] {
  if (revision === null) {
    return stack
  }
  return stack.filter((entry) => entry !== revision)
}

export function deriveWorkbookActorHistoryState(input: {
  readonly actorUserId: string
  readonly rows: readonly WorkbookHistoryStateRow[]
}): WorkbookActorHistoryState {
  const ownRows = [...input.rows]
    .filter((row) => row.actorUserId === input.actorUserId && row.undoBundleJson !== null)
    .toSorted((left, right) => left.revision - right.revision)

  let undoStack: number[] = []
  let redoStack: number[] = []

  for (const row of ownRows) {
    switch (row.eventKind) {
      case 'revertChange': {
        undoStack = removeRevision(undoStack, row.revertsRevision)
        if (row.revertedByRevision === null) {
          redoStack = pushUnique(redoStack, row.revision)
        } else {
          redoStack = removeRevision(redoStack, row.revision)
        }
        break
      }
      case 'redoChange': {
        redoStack = removeRevision(redoStack, row.revertsRevision)
        if (row.revertedByRevision === null) {
          undoStack = pushUnique(undoStack, row.revision)
        } else {
          undoStack = removeRevision(undoStack, row.revision)
        }
        break
      }
      default: {
        redoStack = []
        if (row.revertedByRevision === null) {
          undoStack = pushUnique(undoStack, row.revision)
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
    undoRevision: undoStack.at(-1) ?? null,
    redoRevision: redoStack.at(-1) ?? null,
    undoStack,
    redoStack,
  }
}
