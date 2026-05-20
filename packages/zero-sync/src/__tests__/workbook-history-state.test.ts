import { describe, expect, it } from 'vitest'
import { deriveWorkbookActorHistoryState } from '../workbook-history-state.js'

describe('deriveWorkbookActorHistoryState', () => {
  it('removes overlapping undo and redo entries when another actor changes the same range', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 51, eventKind: 'setCellValue', address: 'A1' }),
        historyRow({
          revision: 52,
          eventKind: 'revertChange',
          address: 'B1',
          revertedByRevision: null,
          revertsRevision: 50,
        }),
        historyRow({
          revision: 53,
          actorUserId: 'morgan@example.com',
          eventKind: 'setCellValue',
          address: 'A1',
        }),
        historyRow({
          revision: 54,
          actorUserId: 'morgan@example.com',
          eventKind: 'setCellValue',
          address: 'B1',
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.canRedo).toBe(false)
    expect(state.undoStack).toEqual([])
    expect(state.redoStack).toEqual([])
  })

  it('preserves actor history when another actor changes a disjoint range', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 61, eventKind: 'setCellValue', address: 'A1' }),
        historyRow({
          revision: 62,
          eventKind: 'revertChange',
          address: 'B1',
          revertedByRevision: null,
          revertsRevision: 60,
        }),
        historyRow({
          revision: 63,
          actorUserId: 'morgan@example.com',
          eventKind: 'setCellValue',
          address: 'C1',
        }),
      ],
    })

    expect(state.canUndo).toBe(true)
    expect(state.canRedo).toBe(true)
    expect(state.undoRevision).toBe(61)
    expect(state.redoRevision).toBe(62)
  })

  it('treats row-scoped history as overlapping cells outside column A on the same rows', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 71, eventKind: 'setCellValue', address: 'B3' }),
        historyRow({
          revision: 72,
          actorUserId: 'morgan@example.com',
          eventKind: 'insertRows',
          address: 'A3',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'rows' },
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.undoStack).toEqual([])
  })

  it('treats column-scoped history as overlapping cells below row one in the same columns', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 81, eventKind: 'setCellValue', address: 'C9' }),
        historyRow({
          revision: 82,
          actorUserId: 'morgan@example.com',
          eventKind: 'deleteColumns',
          address: 'C1',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'D1', scope: 'columns' },
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.undoStack).toEqual([])
  })

  it('keeps structural changes disjoint across row intervals and sheets', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 91, eventKind: 'setCellValue', address: 'B3' }),
        historyRow({
          revision: 92,
          actorUserId: 'morgan@example.com',
          eventKind: 'insertRows',
          address: 'A8',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A8', endAddress: 'A9', scope: 'rows' },
        }),
        historyRow({
          revision: 93,
          actorUserId: 'morgan@example.com',
          eventKind: 'insertRows',
          address: 'A3',
          sheetName: 'Sheet2',
          rangeJson: { sheetName: 'Sheet2', startAddress: 'A3', endAddress: 'A4', scope: 'rows' },
        }),
      ],
    })

    expect(state.canUndo).toBe(true)
    expect(state.undoRevision).toBe(91)
  })

  it('treats malformed structural scope as unknown and invalidates stale history conservatively', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 96, eventKind: 'setCellValue', address: 'Z99' }),
        historyRow({
          revision: 97,
          actorUserId: 'morgan@example.com',
          eventKind: 'insertRows',
          address: 'A3',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.undoStack).toEqual([])
  })

  it('treats malformed persisted range addresses as unknown and invalidates stale history conservatively', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 100, eventKind: 'setCellValue', address: 'Z99' }),
        historyRow({
          revision: 101,
          actorUserId: 'morgan@example.com',
          eventKind: 'setCellValue',
          address: 'A1',
          rangeJson: { sheetName: 'Sheet1', startAddress: 'A0', endAddress: 'A1' },
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.undoStack).toEqual([])
  })

  it('does not fall back to anchor metadata after a range was already marked invalid', () => {
    const state = deriveWorkbookActorHistoryState({
      actorUserId: 'alex@example.com',
      rows: [
        historyRow({ revision: 98, eventKind: 'setCellValue', address: 'Z99' }),
        historyRow({
          revision: 99,
          actorUserId: 'morgan@example.com',
          eventKind: 'insertRows',
          address: 'A3',
          rangeJson: null,
          rangeJsonInvalid: true,
        }),
      ],
    })

    expect(state.canUndo).toBe(false)
    expect(state.undoStack).toEqual([])
  })

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

function historyRow(input: {
  readonly revision: number
  readonly eventKind: string
  readonly actorUserId?: string
  readonly address: string
  readonly sheetName?: string
  readonly rangeJson?: {
    readonly sheetName: string
    readonly startAddress: string
    readonly endAddress: string
    readonly scope?: string
  } | null
  readonly rangeJsonInvalid?: boolean
  readonly revertedByRevision?: number | null
  readonly revertsRevision?: number | null
}) {
  const sheetName = input.sheetName ?? 'Sheet1'
  return {
    revision: input.revision,
    actorUserId: input.actorUserId ?? 'alex@example.com',
    eventKind: input.eventKind,
    undoBundleJson: { kind: 'engineOps', ops: [] },
    revertedByRevision: input.revertedByRevision ?? null,
    revertsRevision: input.revertsRevision ?? null,
    sheetName,
    anchorAddress: input.address,
    rangeJson:
      input.rangeJson === undefined
        ? {
            sheetName,
            startAddress: input.address,
            endAddress: input.address,
          }
        : input.rangeJson,
    rangeJsonInvalid: input.rangeJsonInvalid ?? false,
  } as const
}
