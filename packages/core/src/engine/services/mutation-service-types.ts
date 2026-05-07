import type { Effect } from 'effect'
import type { CellRangeRef, LiteralInput } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import type {
  EngineCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import type { CsvParseOptions } from '../../csv.js'
import type { CommitOp, TransactionRecord } from '../runtime-state.js'
import type { EngineMutationError } from '../errors.js'

export interface EngineMutationService {
  readonly executeTransactionNow: (record: TransactionRecord, source: 'local' | 'restore' | 'undo' | 'redo') => void
  readonly executeTransaction: (
    record: TransactionRecord,
    source: 'local' | 'restore' | 'undo' | 'redo',
  ) => Effect.Effect<void, EngineMutationError>
  readonly executeLocalNow: (
    ops: EngineOp[],
    potentialNewCells?: number,
    options?: { readonly returnUndoOps?: boolean },
  ) => readonly EngineOp[] | null
  readonly executeLocalCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
    options?: {
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ) => readonly EngineOp[] | null
  readonly executeLocalExistingNumericCellMutationAtNow: (
    request: EngineExistingNumericCellMutationRef,
    options?: {
      returnUndoOps?: boolean
    },
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ) => readonly EngineOp[] | null
  readonly applyCellMutationsAt: (
    refs: readonly EngineCellMutationRef[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>
  readonly executeLocal: (
    ops: EngineOp[],
    potentialNewCells?: number,
    options?: { readonly returnUndoOps?: boolean },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>
  readonly applyOpsNow: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      trusted?: boolean
    },
  ) => readonly EngineOp[] | null
  readonly applyOps: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      trusted?: boolean
    },
  ) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>
  readonly captureUndoOps: <Result>(mutate: () => Result) => Effect.Effect<
    {
      result: Result
      undoOps: readonly EngineOp[] | null
    },
    EngineMutationError
  >
  readonly setRangeValues: (range: CellRangeRef, values: readonly (readonly LiteralInput[])[]) => Effect.Effect<void, EngineMutationError>
  readonly setRangeFormulas: (range: CellRangeRef, formulas: readonly (readonly string[])[]) => Effect.Effect<void, EngineMutationError>
  readonly clearRange: (range: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly fillRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly copyRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly moveRange: (source: CellRangeRef, target: CellRangeRef) => Effect.Effect<void, EngineMutationError>
  readonly importSheetCsv: (sheetName: string, csv: string, options?: CsvParseOptions) => Effect.Effect<void, EngineMutationError>
  readonly renderCommit: (ops: CommitOp[]) => Effect.Effect<void, EngineMutationError>
}
