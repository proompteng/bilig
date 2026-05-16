import { formatAddress } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import {
  cellMutationRefToEngineOp,
  type EngineCellMutationAt,
  type EngineCellMutationRef,
  type EngineExistingNumericCellMutationRef,
} from '../../cell-mutations-at.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { PreparedCellAddress, TransactionRecord } from '../runtime-state.js'

type OpsTransactionRecord = Extract<TransactionRecord, { kind: 'ops' }>

export function createOpsTransactionRecord(
  ops: EngineOp[],
  potentialNewCells?: number,
  preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
): OpsTransactionRecord {
  const record: OpsTransactionRecord = { kind: 'ops', ops }
  if (potentialNewCells !== undefined) {
    record.potentialNewCells = potentialNewCells
  }
  if (preparedCellAddressesByOpIndex !== undefined) {
    record.preparedCellAddressesByOpIndex = preparedCellAddressesByOpIndex
  }
  return record
}

export function createLazySingleOpTransactionRecord(op: EngineOp, potentialNewCells?: number): TransactionRecord {
  return potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells }
}

export function createLazyCellMutationTransactionRecord(
  refs: readonly EngineCellMutationRef[],
  potentialNewCells?: number,
): TransactionRecord {
  return potentialNewCells === undefined ? { kind: 'cell-mutations', refs } : { kind: 'cell-mutations', refs, potentialNewCells }
}

export function createSingleExistingNumericCellMutationTransactionRecord(
  request: EngineExistingNumericCellMutationRef,
  potentialNewCells: number,
): TransactionRecord {
  return {
    kind: 'single-existing-numeric-cell-mutation',
    sheetId: request.sheetId,
    row: request.row,
    col: request.col,
    cellIndex: request.cellIndex,
    value: request.value,
    potentialNewCells,
  }
}

export function singleExistingNumericCellMutationRecordToRef(
  record: Extract<TransactionRecord, { kind: 'single-existing-numeric-cell-mutation' }>,
): EngineCellMutationRef {
  return {
    sheetId: record.sheetId,
    cellIndex: record.cellIndex,
    mutation: {
      kind: 'setCellValue',
      row: record.row,
      col: record.col,
      value: record.value,
    },
  }
}

export function createLazyMaterializedCellMutationTransactionRecord(
  materializeRefs: () => EngineCellMutationRef[],
  potentialNewCells?: number,
): TransactionRecord {
  let cachedRefs: EngineCellMutationRef[] | undefined
  const record: { kind: 'cell-mutations'; refs: readonly EngineCellMutationRef[]; potentialNewCells?: number } = {
    kind: 'cell-mutations',
    get refs() {
      cachedRefs ??= materializeRefs()
      return cachedRefs
    },
  }
  if (potentialNewCells !== undefined) {
    record.potentialNewCells = potentialNewCells
  }
  return record
}

export interface RenderCommitCellMutation {
  readonly sheetName: string
  readonly mutation: EngineCellMutationAt
}

export function renderCommitCellMutationToEngineOp(entry: RenderCommitCellMutation): EngineOp {
  const address = formatAddress(entry.mutation.row, entry.mutation.col)
  switch (entry.mutation.kind) {
    case 'setCellValue':
      return {
        kind: 'setCellValue',
        sheetName: entry.sheetName,
        address,
        value: entry.mutation.value,
      }
    case 'setCellFormula':
      return {
        kind: 'setCellFormula',
        sheetName: entry.sheetName,
        address,
        formula: entry.mutation.formula,
      }
    case 'clearCell':
      return {
        kind: 'clearCell',
        sheetName: entry.sheetName,
        address,
      }
  }
}

export function createLazyRenderCommitTransactionRecord(
  prefixOps: readonly EngineOp[],
  cellMutations: readonly RenderCommitCellMutation[],
  potentialNewCells?: number,
): TransactionRecord {
  let cachedOps: EngineOp[] | undefined
  const record: { kind: 'ops'; ops: EngineOp[]; potentialNewCells?: number } = {
    kind: 'ops',
    get ops() {
      cachedOps ??= [
        ...prefixOps.map((op) => structuredClone(op)),
        ...cellMutations.map((entry) => renderCommitCellMutationToEngineOp(entry)),
      ]
      return cachedOps
    },
  }
  if (potentialNewCells !== undefined) {
    record.potentialNewCells = potentialNewCells
  }
  return record
}

export function transactionRecordOps(workbook: WorkbookStore, record: TransactionRecord): readonly EngineOp[] {
  if (record.kind === 'single-op') {
    return [record.op]
  }
  if (record.kind === 'single-existing-numeric-cell-mutation') {
    return [cellMutationRefToEngineOp(workbook, singleExistingNumericCellMutationRecordToRef(record))]
  }
  if (record.kind === 'cell-mutations') {
    return record.refs.map((ref) => cellMutationRefToEngineOp(workbook, ref))
  }
  return record.ops
}

export function cloneTransactionRecordOps(workbook: WorkbookStore, record: TransactionRecord): EngineOp[] {
  if (record.kind === 'single-op') {
    return [structuredClone(record.op)]
  }
  if (record.kind === 'single-existing-numeric-cell-mutation') {
    return [cellMutationRefToEngineOp(workbook, singleExistingNumericCellMutationRecordToRef(record))]
  }
  if (record.kind === 'cell-mutations') {
    return record.refs.map((ref) => cellMutationRefToEngineOp(workbook, ref))
  }
  return structuredClone(record.ops)
}
