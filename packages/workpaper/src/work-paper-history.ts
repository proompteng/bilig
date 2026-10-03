import { formatAddress } from '@bilig/formula'
import type { RawCellContent } from './work-paper-types.js'

export type WorkPaperHistoryTransactionRecord =
  | { kind: 'ops'; ops: unknown[]; potentialNewCells?: number }
  | { kind: 'single-op'; op: unknown; potentialNewCells?: number }
  | { kind: 'rename-sheet'; oldName: string; newName: string; potentialNewCells?: number }
  | {
      kind: 'single-existing-numeric-cell-mutation'
      sheetId: number
      row: number
      col: number
      cellIndex: number
      value: number
      potentialNewCells?: number
    }
  | {
      kind: 'existing-numeric-cell-mutations'
      sheetIds: Uint32Array
      cellIndexPlusOnes: Uint32Array
      rows: Uint32Array
      cols: Uint32Array
      numbers: Float64Array
      potentialNewCells?: number
    }
  | {
      kind: 'cell-mutations'
      refs: Array<{
        sheetId: number
        mutation:
          | { kind: 'setCellValue'; row: number; col: number; value: RawCellContent }
          | { kind: 'setCellFormula'; row: number; col: number; formula: string }
          | { kind: 'clearCell'; row: number; col: number }
      }>
      potentialNewCells?: number
    }

export interface WorkPaperHistoryRecord {
  forward: WorkPaperHistoryTransactionRecord
  inverse: WorkPaperHistoryTransactionRecord
}

export type WorkPaperSheetNameResolver = (sheetId: number) => string | undefined

function isWorkPaperHistoryRecordArray(value: unknown): value is WorkPaperHistoryRecord[] {
  return Array.isArray(value)
}

export function readWorkPaperHistoryStack(owner: object, propertyName: 'undoStack' | 'redoStack'): WorkPaperHistoryRecord[] {
  const stack = Reflect.get(owner, propertyName)
  if (!isWorkPaperHistoryRecordArray(stack)) {
    return []
  }
  return stack
}

export function workPaperHistoryTopIsCellMutations(stack: readonly WorkPaperHistoryRecord[]): boolean {
  const kind = stack.at(-1)?.forward.kind
  return kind === 'cell-mutations' || kind === 'single-existing-numeric-cell-mutation' || kind === 'existing-numeric-cell-mutations'
}

export function clearWorkPaperHistoryStacks(undoStack: WorkPaperHistoryRecord[], redoStack: WorkPaperHistoryRecord[]): void {
  undoStack.length = 0
  redoStack.length = 0
}

export function cloneWorkPaperHistoryRecords(records: readonly WorkPaperHistoryRecord[]): WorkPaperHistoryRecord[] {
  return records.map((record) => ({
    forward: cloneWorkPaperHistoryTransactionRecord(record.forward),
    inverse: cloneWorkPaperHistoryTransactionRecord(record.inverse),
  }))
}

function cloneWorkPaperHistoryTransactionRecord(record: WorkPaperHistoryTransactionRecord): WorkPaperHistoryTransactionRecord {
  switch (record.kind) {
    case 'ops':
      return {
        kind: 'ops',
        ops: record.ops.map((op) => structuredClone(op)),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'single-op':
      return {
        kind: 'single-op',
        op: structuredClone(record.op),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'rename-sheet':
      return {
        kind: 'rename-sheet',
        oldName: record.oldName,
        newName: record.newName,
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'single-existing-numeric-cell-mutation':
      return {
        kind: 'single-existing-numeric-cell-mutation',
        sheetId: record.sheetId,
        row: record.row,
        col: record.col,
        cellIndex: record.cellIndex,
        value: record.value,
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'existing-numeric-cell-mutations':
      return {
        kind: 'existing-numeric-cell-mutations',
        sheetIds: record.sheetIds.slice(),
        cellIndexPlusOnes: record.cellIndexPlusOnes.slice(),
        rows: record.rows.slice(),
        cols: record.cols.slice(),
        numbers: record.numbers.slice(),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
    case 'cell-mutations':
      return {
        kind: 'cell-mutations',
        refs: record.refs.map((ref) => ({
          sheetId: ref.sheetId,
          mutation: { ...ref.mutation },
        })),
        ...(record.potentialNewCells !== undefined ? { potentialNewCells: record.potentialNewCells } : {}),
      }
  }
}

export function workPaperHistoryTransactionOps(
  record: WorkPaperHistoryTransactionRecord,
  resolveSheetName: WorkPaperSheetNameResolver,
): unknown[] {
  switch (record.kind) {
    case 'ops':
      return record.ops
    case 'single-op':
      return [record.op]
    case 'rename-sheet':
      return [{ kind: 'renameSheet', oldName: record.oldName, newName: record.newName }]
    case 'single-existing-numeric-cell-mutation': {
      const sheetName = resolveSheetName(record.sheetId)
      return sheetName
        ? [
            {
              kind: 'setCellValue',
              sheetName,
              address: formatAddress(record.row, record.col),
              value: record.value,
            },
          ]
        : []
    }
    case 'existing-numeric-cell-mutations': {
      const ops: unknown[] = []
      for (let index = 0; index < record.sheetIds.length; index += 1) {
        const sheetName = resolveSheetName(record.sheetIds[index]!)
        if (sheetName) {
          ops.push({
            kind: 'setCellValue',
            sheetName,
            address: formatAddress(record.rows[index]!, record.cols[index]!),
            value: record.numbers[index]!,
          })
        }
      }
      return ops
    }
    case 'cell-mutations':
      return record.refs.flatMap((ref) => {
        const sheetName = resolveSheetName(ref.sheetId)
        if (!sheetName) {
          return []
        }
        const address = formatAddress(ref.mutation.row, ref.mutation.col)
        switch (ref.mutation.kind) {
          case 'setCellValue':
            return [
              {
                kind: 'setCellValue',
                sheetName,
                address,
                value: ref.mutation.value,
              },
            ]
          case 'setCellFormula':
            return [
              {
                kind: 'setCellFormula',
                sheetName,
                address,
                formula: ref.mutation.formula,
              },
            ]
          case 'clearCell':
            return [
              {
                kind: 'clearCell',
                sheetName,
                address,
              },
            ]
        }
      })
  }
}

export function tryMergeTypedCellMutationHistory(entries: readonly WorkPaperHistoryRecord[]): WorkPaperHistoryRecord | null {
  if (entries.length === 0 || entries.some((entry) => entry.forward.kind !== 'cell-mutations' || entry.inverse.kind !== 'cell-mutations')) {
    return null
  }
  const forward: WorkPaperHistoryTransactionRecord = {
    kind: 'cell-mutations',
    refs: entries.flatMap((entry) => (entry.forward.kind === 'cell-mutations' ? entry.forward.refs : [])),
  }
  const forwardPotentialNewCells = sumNumbers(entries.map((entry) => entry.forward.potentialNewCells))
  if (forwardPotentialNewCells !== undefined) {
    forward.potentialNewCells = forwardPotentialNewCells
  }
  const inverse: WorkPaperHistoryTransactionRecord = {
    kind: 'cell-mutations',
    refs: entries.toReversed().flatMap((entry) => (entry.inverse.kind === 'cell-mutations' ? entry.inverse.refs : [])),
  }
  const inversePotentialNewCells = sumNumbers(entries.map((entry) => entry.inverse.potentialNewCells))
  if (inversePotentialNewCells !== undefined) {
    inverse.potentialNewCells = inversePotentialNewCells
  }
  return {
    forward,
    inverse,
  }
}

export function mergeWorkPaperUndoHistory(
  undoStack: WorkPaperHistoryRecord[],
  startIndex: number,
  resolveSheetName: WorkPaperSheetNameResolver,
): void {
  if (undoStack.length - startIndex <= 1) {
    return
  }
  const entries = undoStack.splice(startIndex)
  const merged = tryMergeTypedCellMutationHistory(entries) ?? mergeWorkPaperOperationHistory(entries, resolveSheetName)
  undoStack.push(merged)
}

function mergeWorkPaperOperationHistory(
  entries: readonly WorkPaperHistoryRecord[],
  resolveSheetName: WorkPaperSheetNameResolver,
): WorkPaperHistoryRecord {
  const forward = lazyMergedOperationRecord(entries, (entry) => entry.forward, resolveSheetName)
  const forwardPotentialNewCells = sumNumbers(entries.map((entry) => entry.forward.potentialNewCells))
  if (forwardPotentialNewCells !== undefined) {
    forward.potentialNewCells = forwardPotentialNewCells
  }
  const inverse = lazyMergedOperationRecord(entries.toReversed(), (entry) => entry.inverse, resolveSheetName)
  const inversePotentialNewCells = sumNumbers(entries.map((entry) => entry.inverse.potentialNewCells))
  if (inversePotentialNewCells !== undefined) {
    inverse.potentialNewCells = inversePotentialNewCells
  }
  return { forward, inverse }
}

function lazyMergedOperationRecord(
  entries: readonly WorkPaperHistoryRecord[],
  selectRecord: (entry: WorkPaperHistoryRecord) => WorkPaperHistoryTransactionRecord,
  resolveSheetName: WorkPaperSheetNameResolver,
): Extract<WorkPaperHistoryTransactionRecord, { kind: 'ops' }> {
  let cachedOps: unknown[] | undefined
  return {
    kind: 'ops',
    get ops() {
      cachedOps ??= entries.flatMap((entry) => workPaperHistoryTransactionOps(selectRecord(entry), resolveSheetName))
      return cachedOps
    },
  }
}

function sumNumbers(values: Array<number | undefined>): number | undefined {
  const filtered = values.filter((value): value is number => typeof value === 'number')
  if (filtered.length === 0) {
    return undefined
  }
  return filtered.reduce((sum, value) => sum + value, 0)
}
