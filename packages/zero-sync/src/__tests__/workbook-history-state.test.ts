import { describe, expect, it } from 'vitest'
import { deriveWorkbookActorHistoryState } from '../workbook-history-state.js'

describe('deriveWorkbookActorHistoryState', () => {
  it('preserves older redo entries after a newer redo is applied', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        {
          revision: 21,
          actorUserId: 'alex@example.com',
          eventKind: 'setCellValue',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: 24,
          revertsRevision: null,
        },
        {
          revision: 22,
          actorUserId: 'alex@example.com',
          eventKind: 'setCellValue',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: 23,
          revertsRevision: null,
        },
        {
          revision: 23,
          actorUserId: 'alex@example.com',
          eventKind: 'revertChange',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: null,
          revertsRevision: 22,
        },
        {
          revision: 24,
          actorUserId: 'alex@example.com',
          eventKind: 'revertChange',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: 25,
          revertsRevision: 21,
        },
        {
          revision: 25,
          actorUserId: 'alex@example.com',
          eventKind: 'redoChange',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: null,
          revertsRevision: 24,
        },
      ],
    })

    expect(state.canUndo).toBe(true)
    expect(state.canRedo).toBe(true)
    expect(state.undoRevision).toBe(25)
    expect(state.redoRevision).toBe(23)
  })

  it('clears redo after a fresh authored change is appended after an undo', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        {
          revision: 31,
          actorUserId: 'alex@example.com',
          eventKind: 'setCellValue',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: 32,
          revertsRevision: null,
        },
        {
          revision: 32,
          actorUserId: 'alex@example.com',
          eventKind: 'revertChange',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: null,
          revertsRevision: 31,
        },
        {
          revision: 33,
          actorUserId: 'alex@example.com',
          eventKind: 'setCellValue',
          undoBundleJson: { kind: 'engineOps', ops: [] },
          revertedByRevision: null,
          revertsRevision: null,
        },
      ],
    })

    expect(state.canUndo).toBe(true)
    expect(state.canRedo).toBe(false)
    expect(state.undoRevision).toBe(33)
    expect(state.redoRevision).toBeNull()
  })

  it('handles a larger history with multiple remaining redo levels', () => {
    const rows = [
      {
        revision: 41,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: null,
      },
      {
        revision: 42,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: null,
      },
      {
        revision: 43,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: null,
      },
      {
        revision: 44,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: 47,
        revertsRevision: null,
      },
      {
        revision: 45,
        actorUserId: 'alex@example.com',
        eventKind: 'setCellValue',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: 46,
        revertsRevision: null,
      },
      {
        revision: 46,
        actorUserId: 'alex@example.com',
        eventKind: 'revertChange',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: 45,
      },
      {
        revision: 47,
        actorUserId: 'alex@example.com',
        eventKind: 'revertChange',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: 48,
        revertsRevision: 44,
      },
      {
        revision: 48,
        actorUserId: 'alex@example.com',
        eventKind: 'redoChange',
        undoBundleJson: { kind: 'engineOps', ops: [] },
        revertedByRevision: null,
        revertsRevision: 47,
      },
    ] as const

    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows,
    })

    expect(state.canUndo).toBe(true)
    expect(state.canRedo).toBe(true)
    expect(state.undoRevision).toBe(48)
    expect(state.redoRevision).toBe(46)
    expect(state.redoStack).toEqual([46])
  })
})
