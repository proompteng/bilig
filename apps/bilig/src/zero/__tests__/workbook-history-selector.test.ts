import { describe, expect, it } from 'vitest'
import { selectLatestRedoableWorkbookChangeRevision, selectLatestUndoableWorkbookChangeRevision } from '../workbook-history-selector.js'

describe('workbook-history-selector', () => {
  it('keeps the next older redo revision available after a newer redo succeeds', () => {
    const rows = [
      {
        revision: 21,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: 24,
        revertsRevision: null,
      },
      {
        revision: 22,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: 23,
        revertsRevision: null,
      },
      {
        revision: 23,
        actorUserId: 'alex@example.com',
        eventKind: 'revertChange',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: 22,
      },
      {
        revision: 24,
        actorUserId: 'alex@example.com',
        eventKind: 'revertChange',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: 25,
        revertsRevision: 21,
      },
      {
        revision: 25,
        actorUserId: 'alex@example.com',
        eventKind: 'redoChange',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: 24,
      },
    ] as const

    expect(
      selectLatestUndoableWorkbookChangeRevision({
        actorUserId: 'alex@example.com',
        rows,
      }),
    ).toBe(25)
    expect(
      selectLatestRedoableWorkbookChangeRevision({
        actorUserId: 'alex@example.com',
        rows,
      }),
    ).toBe(23)
  })

  it('clears redo after a fresh authored branch edit', () => {
    const rows = [
      {
        revision: 31,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: 32,
        revertsRevision: null,
      },
      {
        revision: 32,
        actorUserId: 'alex@example.com',
        eventKind: 'revertChange',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: 31,
      },
      {
        revision: 33,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundle: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: null,
      },
    ] as const

    expect(
      selectLatestUndoableWorkbookChangeRevision({
        actorUserId: 'alex@example.com',
        rows,
      }),
    ).toBe(33)
    expect(
      selectLatestRedoableWorkbookChangeRevision({
        actorUserId: 'alex@example.com',
        rows,
      }),
    ).toBeNull()
  })
})
