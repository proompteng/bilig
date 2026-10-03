import { describe, expect, it } from 'vitest'
import {
  clearWorkPaperHistoryStacks,
  cloneWorkPaperHistoryRecords,
  mergeWorkPaperUndoHistory,
  readWorkPaperHistoryStack,
  tryMergeTypedCellMutationHistory,
  workPaperHistoryTopIsCellMutations,
  workPaperHistoryTransactionOps,
  type WorkPaperHistoryRecord,
} from '../work-paper-history.js'

function cellMutationRecord(sheetId: number, row: number, value: number): WorkPaperHistoryRecord {
  return {
    forward: {
      kind: 'cell-mutations',
      refs: [{ sheetId, mutation: { kind: 'setCellValue', row, col: 0, value } }],
      potentialNewCells: 1,
    },
    inverse: {
      kind: 'cell-mutations',
      refs: [{ sheetId, mutation: { kind: 'clearCell', row, col: 0 } }],
      potentialNewCells: 0,
    },
  }
}

describe('work paper history helpers', () => {
  it('reads and clears reflected engine stacks', () => {
    const owner = {
      undoStack: [cellMutationRecord(1, 0, 1)],
      redoStack: [cellMutationRecord(1, 1, 2)],
    }

    const undoStack = readWorkPaperHistoryStack(owner, 'undoStack')
    const redoStack = readWorkPaperHistoryStack(owner, 'redoStack')

    expect(undoStack).toHaveLength(1)
    expect(redoStack).toHaveLength(1)
    clearWorkPaperHistoryStacks(undoStack, redoStack)
    expect(owner.undoStack).toEqual([])
    expect(owner.redoStack).toEqual([])
    expect(readWorkPaperHistoryStack({}, 'undoStack')).toEqual([])
  })

  it('detects typed cell mutation records at the top of a stack', () => {
    expect(workPaperHistoryTopIsCellMutations([])).toBe(false)
    expect(workPaperHistoryTopIsCellMutations([{ forward: { kind: 'single-op', op: {} }, inverse: { kind: 'single-op', op: {} } }])).toBe(
      false,
    )
    expect(workPaperHistoryTopIsCellMutations([cellMutationRecord(1, 0, 1)])).toBe(true)
    expect(
      workPaperHistoryTopIsCellMutations([
        {
          forward: {
            kind: 'single-existing-numeric-cell-mutation',
            sheetId: 1,
            row: 0,
            col: 0,
            cellIndex: 1,
            value: 4,
          },
          inverse: { kind: 'single-op', op: {} },
        },
      ]),
    ).toBe(true)
  })

  it('materializes typed transactions into engine operations', () => {
    expect(workPaperHistoryTransactionOps({ kind: 'single-op', op: { kind: 'noop' } }, () => 'Sheet1')).toEqual([{ kind: 'noop' }])
    expect(
      workPaperHistoryTransactionOps(
        {
          kind: 'single-existing-numeric-cell-mutation',
          sheetId: 2,
          row: 4,
          col: 1,
          cellIndex: 10,
          value: 8,
        },
        (sheetId) => (sheetId === 2 ? 'Data' : undefined),
      ),
    ).toEqual([{ kind: 'setCellValue', sheetName: 'Data', address: 'B5', value: 8 }])
    expect(workPaperHistoryTransactionOps(cellMutationRecord(2, 1, 9).forward, () => undefined)).toEqual([])
  })

  it('deep clones history records used for transaction rollback snapshots', () => {
    const source: WorkPaperHistoryRecord[] = [
      {
        forward: {
          kind: 'ops',
          ops: [{ kind: 'setCellValue', nested: { value: 1 } }],
          potentialNewCells: 1,
        },
        inverse: {
          kind: 'single-op',
          op: { kind: 'clearCell', nested: { address: 'A1' } },
        },
      },
      cellMutationRecord(2, 3, 9),
    ]

    const cloned = cloneWorkPaperHistoryRecords(source)
    if (cloned[0]?.forward.kind !== 'ops' || source[0]?.forward.kind !== 'ops') {
      throw new Error('expected generic ops history')
    }
    if (cloned[0].inverse.kind !== 'single-op' || source[0].inverse.kind !== 'single-op') {
      throw new Error('expected single-op inverse history')
    }
    if (cloned[1]?.forward.kind !== 'cell-mutations' || source[1]?.forward.kind !== 'cell-mutations') {
      throw new Error('expected typed cell-mutation history')
    }

    Reflect.set(Reflect.get(cloned[0].forward.ops[0], 'nested'), 'value', 99)
    Reflect.set(Reflect.get(cloned[0].inverse.op, 'nested'), 'address', 'Z9')
    cloned[1].forward.refs[0].mutation.row = 99

    expect(Reflect.get(Reflect.get(source[0].forward.ops[0], 'nested'), 'value')).toBe(1)
    expect(Reflect.get(Reflect.get(source[0].inverse.op, 'nested'), 'address')).toBe('A1')
    expect(source[1].forward.refs[0]?.mutation.row).toBe(3)
    expect(cloned).not.toEqual(cloneWorkPaperHistoryRecords(source))
    expect(cloned[0].forward.ops[0]).not.toBe(source[0].forward.ops[0])
    expect(cloned[1].forward.refs[0]).not.toBe(source[1].forward.refs[0])
  })

  it('merges typed cell mutation history without degrading to generic ops', () => {
    const merged = tryMergeTypedCellMutationHistory([cellMutationRecord(1, 0, 1), cellMutationRecord(1, 1, 2)])

    expect(merged?.forward.kind).toBe('cell-mutations')
    expect(merged?.inverse.kind).toBe('cell-mutations')
    expect(merged?.forward.potentialNewCells).toBe(2)
    expect(merged?.inverse.potentialNewCells).toBe(0)
    expect(merged?.forward.kind === 'cell-mutations' ? merged.forward.refs.map((ref) => ref.mutation.row) : []).toEqual([0, 1])
    expect(merged?.inverse.kind === 'cell-mutations' ? merged.inverse.refs.map((ref) => ref.mutation.row) : []).toEqual([1, 0])
  })

  it('merges mixed history into generic ops', () => {
    const undoStack: WorkPaperHistoryRecord[] = [
      { forward: { kind: 'single-op', op: { kind: 'a' }, potentialNewCells: 1 }, inverse: { kind: 'single-op', op: { kind: 'undo-a' } } },
      cellMutationRecord(1, 2, 6),
    ]

    mergeWorkPaperUndoHistory(undoStack, 0, () => 'Sheet1')

    expect(undoStack).toHaveLength(1)
    const merged = undoStack[0]
    expect(merged.forward.kind).toBe('ops')
    expect(merged.inverse.kind).toBe('ops')
    expect(merged.forward.potentialNewCells).toBe(2)
    expect(merged.forward.kind === 'ops' ? merged.forward.ops : []).toEqual([
      { kind: 'a' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A3', value: 6 },
    ])
  })

  it('defers mixed-history op materialization until undo or inspection reads it', () => {
    const undoStack: WorkPaperHistoryRecord[] = [
      { forward: { kind: 'single-op', op: { kind: 'a' }, potentialNewCells: 1 }, inverse: { kind: 'single-op', op: { kind: 'undo-a' } } },
      cellMutationRecord(1, 2, 6),
    ]
    let resolveCalls = 0

    mergeWorkPaperUndoHistory(undoStack, 0, () => {
      resolveCalls += 1
      return 'Sheet1'
    })

    expect(resolveCalls).toBe(0)
    const merged = undoStack[0]
    expect(merged.forward.kind).toBe('ops')
    expect(resolveCalls).toBe(0)
    expect(merged.forward.kind === 'ops' ? merged.forward.ops : []).toEqual([
      { kind: 'a' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A3', value: 6 },
    ])
    expect(resolveCalls).toBe(1)
  })
})
