import { describe, expect, it } from 'vitest'
import type { EngineOp } from '@bilig/workbook-domain'
import { WorkbookStore } from '../workbook-store.js'
import {
  cloneTransactionRecordOps,
  createLazyCellMutationTransactionRecord,
  createLazyMaterializedCellMutationTransactionRecord,
  createLazyRenderCommitTransactionRecord,
  createLazySingleOpTransactionRecord,
  createOpsTransactionRecord,
  createSingleExistingNumericCellMutationTransactionRecord,
  singleExistingNumericCellMutationRecordToRef,
  transactionRecordOps,
} from '../engine/services/mutation-transaction-records.js'

describe('mutation transaction records', () => {
  it('creates compact single-op and cell-mutation transaction records', () => {
    const op: EngineOp = { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 }
    const singleOp = createLazySingleOpTransactionRecord(op, 1)
    const cellMutation = createLazyCellMutationTransactionRecord(
      [{ sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1+1' } }],
      0,
    )

    expect(singleOp).toEqual({ kind: 'single-op', op, potentialNewCells: 1 })
    expect(cellMutation).toEqual({
      kind: 'cell-mutations',
      refs: [{ sheetId: 1, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1+1' } }],
      potentialNewCells: 0,
    })
  })

  it('creates op transaction records without losing zero potential cells or prepared addresses', () => {
    const ops: EngineOp[] = [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'B1', formula: 'A1*2' },
    ]
    const prepared = [
      { row: 0, col: 0 },
      { row: 0, col: 1 },
    ]

    expect(createOpsTransactionRecord(ops)).toEqual({ kind: 'ops', ops })
    expect(createOpsTransactionRecord(ops, 0, prepared)).toEqual({
      kind: 'ops',
      ops,
      potentialNewCells: 0,
      preparedCellAddressesByOpIndex: prepared,
    })
  })

  it('lazily materializes cell mutation refs once', () => {
    let calls = 0
    const record = createLazyMaterializedCellMutationTransactionRecord(() => {
      calls += 1
      return [{ sheetId: 1, cellIndex: 4, mutation: { kind: 'clearCell', row: 1, col: 2 } }]
    }, 0)

    expect(calls).toBe(0)
    expect(record).toMatchObject({ kind: 'cell-mutations', potentialNewCells: 0 })
    expect(record.refs).toEqual([{ sheetId: 1, cellIndex: 4, mutation: { kind: 'clearCell', row: 1, col: 2 } }])
    expect(record.refs).toEqual([{ sheetId: 1, cellIndex: 4, mutation: { kind: 'clearCell', row: 1, col: 2 } }])
    expect(calls).toBe(1)
  })

  it('materializes existing numeric mutation records into refs and engine ops', () => {
    const workbook = new WorkbookStore('txn')
    const sheet = workbook.createSheet('Sheet1')
    const record = createSingleExistingNumericCellMutationTransactionRecord(
      { sheetId: sheet.id, row: 2, col: 3, cellIndex: 9, value: 42 },
      0,
    )

    expect(singleExistingNumericCellMutationRecordToRef(record)).toEqual({
      sheetId: sheet.id,
      cellIndex: 9,
      mutation: { kind: 'setCellValue', row: 2, col: 3, value: 42 },
    })
    expect(transactionRecordOps(workbook, record)).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D3', value: 42 }])
  })

  it('builds lazy render commit transactions from prefix ops and cell mutations', () => {
    const record = createLazyRenderCommitTransactionRecord(
      [{ kind: 'upsertSheet', name: 'Sheet1', order: 0 }],
      [
        { sheetName: 'Sheet1', mutation: { kind: 'setCellValue', row: 0, col: 0, value: 'x' } },
        { sheetName: 'Sheet1', mutation: { kind: 'setCellFormula', row: 1, col: 0, formula: 'A1' } },
        { sheetName: 'Sheet1', mutation: { kind: 'clearCell', row: 2, col: 0 } },
      ],
      2,
    )

    expect(record).toMatchObject({ kind: 'ops', potentialNewCells: 2 })
    expect(record.ops).toEqual([
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'x' },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A2', formula: 'A1' },
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'A3' },
    ])
  })

  it('clones op-backed transaction records before returning undo ops', () => {
    const workbook = new WorkbookStore('txn')
    const record = {
      kind: 'ops' as const,
      ops: [{ kind: 'setCellValue' as const, sheetName: 'Sheet1', address: 'A1', value: 1 }],
    }

    const cloned = cloneTransactionRecordOps(workbook, record)
    cloned[0] = { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 2 }

    expect(record.ops).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }])
  })
})
