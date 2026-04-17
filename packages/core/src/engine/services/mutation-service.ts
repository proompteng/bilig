import { Effect } from 'effect'
import {
  formatAddress,
  parseCellAddress,
  rewriteAddressForStructuralTransform,
  rewriteRangeForStructuralTransform,
  translateFormulaReferences,
} from '@bilig/formula'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import { ValueTag, type CellRangeRef, type CellSnapshot, type LiteralInput } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import { normalizeRange } from '../../engine-range-utils.js'
import { cloneCellStyleRecord } from '../../engine-style-utils.js'
import { structuralTransformForOp } from '../../engine-structural-utils.js'
import { restoreFormatRangeOps, restoreStyleRangeOps } from '../../engine-range-format-ops.js'
import { sheetMetadataToOps } from '../../engine-snapshot-utils.js'
import { parseCsv, parseCsvCellInput } from '../../csv.js'
import { createBatch } from '../../replica-state.js'
import type { WorkbookStore } from '../../workbook-store.js'
import {
  cellMutationRefToEngineOp,
  cloneCellMutationRef,
  countPotentialNewCellsForMutationRefs,
  type EngineCellMutationRef,
} from '../../cell-mutations-at.js'
import type { CommitOp, EngineRuntimeState, PreparedCellAddress, TransactionRecord } from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { tryBuildFastMutationHistory, type FastMutationHistoryResult } from './mutation-history-fast-path.js'

function mutationErrorMessage(message: string, cause: unknown): string {
  return cause instanceof Error && cause.message.length > 0 ? cause.message : message
}

function getMatrixCell(matrix: readonly (readonly CellSnapshot[])[], rowIndex: number, colIndex: number): CellSnapshot {
  const row = matrix[rowIndex]
  if (row === undefined) {
    throw new RangeError(`Missing source row at index ${rowIndex}`)
  }
  const cell = row[colIndex]
  if (cell === undefined) {
    throw new RangeError(`Missing source cell at row ${rowIndex}, column ${colIndex}`)
  }
  return cell
}

function createLazySingleOpTransactionRecord(op: EngineOp, potentialNewCells?: number): TransactionRecord {
  return potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells }
}

function createLazyCellMutationTransactionRecord(
  workbook: WorkbookStore,
  refs: readonly EngineCellMutationRef[],
  potentialNewCells?: number,
): TransactionRecord {
  const record: { kind: 'ops'; ops: EngineOp[]; potentialNewCells?: number } = {
    kind: 'ops',
    get ops() {
      cachedOps ??= refs.map((ref) => cellMutationRefToEngineOp(workbook, ref))
      return cachedOps
    },
  }
  let cachedOps: EngineOp[] | undefined
  if (potentialNewCells !== undefined) {
    record.potentialNewCells = potentialNewCells
  }
  return record
}

function transactionRecordOps(record: TransactionRecord): readonly EngineOp[] {
  return record.kind === 'single-op' ? [record.op] : record.ops
}

function cloneTransactionRecordOps(record: TransactionRecord): EngineOp[] {
  return record.kind === 'single-op' ? [structuredClone(record.op)] : structuredClone(record.ops)
}

interface ComparableCellState {
  formula?: string
  value: LiteralInput | null
  format: string | null
  authoredBlank?: boolean
}

function translateFormulaForTarget(
  formula: string,
  sourceSheetName: string,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): string {
  if (sourceSheetName !== targetSheetName) {
    return formula
  }
  const source = parseCellAddress(sourceAddress, sourceSheetName)
  const target = parseCellAddress(targetAddress, targetSheetName)
  return translateFormulaReferences(formula, target.row - source.row, target.col - source.col)
}

export interface EngineMutationService {
  readonly executeTransactionNow: (record: TransactionRecord, source: 'local' | 'restore' | 'undo' | 'redo') => void
  readonly executeTransaction: (
    record: TransactionRecord,
    source: 'local' | 'restore' | 'undo' | 'redo',
  ) => Effect.Effect<void, EngineMutationError>
  readonly executeLocalNow: (ops: EngineOp[], potentialNewCells?: number) => readonly EngineOp[] | null
  readonly executeLocalCellMutationsAtNow: (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
    options?: {
      returnUndoOps?: boolean
      reuseRefs?: boolean
    },
  ) => readonly EngineOp[] | null
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
  readonly executeLocal: (ops: EngineOp[], potentialNewCells?: number) => Effect.Effect<readonly EngineOp[] | null, EngineMutationError>
  readonly applyOpsNow: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      trusted?: boolean
    },
  ) => readonly EngineOp[] | null
  readonly applyOps: (
    ops: readonly EngineOp[],
    options?: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
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
  readonly importSheetCsv: (sheetName: string, csv: string) => Effect.Effect<void, EngineMutationError>
  readonly renderCommit: (ops: CommitOp[]) => Effect.Effect<void, EngineMutationError>
}

export function createEngineMutationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'replicaState'
    | 'batchListeners'
    | 'formulas'
    | 'undoStack'
    | 'redoStack'
    | 'trackReplicaVersions'
    | 'getSyncClientConnection'
    | 'getTransactionReplayDepth'
    | 'setTransactionReplayDepth'
  > & {
    readonly workbook: WorkbookStore
  }
  readonly captureSheetCellState: (sheetName: string) => EngineOp[]
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => EngineOp[]
  readonly captureStoredCellOps: (cellIndex: number, sheetName: string, address: string) => EngineOp[]
  readonly restoreCellOps: (sheetName: string, address: string) => EngineOp[]
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot
  readonly readRangeCells: (range: CellRangeRef) => CellSnapshot[][]
  readonly toCellStateOps: (
    sheetName: string,
    address: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly applyBatchNow: (
    batch: EngineOpBatch,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
    preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[],
  ) => void
  readonly applyCellMutationsAtBatchNow: (
    refs: readonly EngineCellMutationRef[],
    batch: EngineOpBatch | null,
    source: 'local' | 'restore' | 'undo' | 'redo',
    potentialNewCells?: number,
  ) => void
}): EngineMutationService {
  const emptyBatchOps: EngineOp[] = []
  const shouldCreateLocalBatch = (): boolean =>
    args.state.trackReplicaVersions || (args.state.batchListeners?.size ?? 0) > 0 || args.state.getSyncClientConnection?.() !== null

  const restoreCellOpFromRef = (ref: EngineCellMutationRef): EngineOp => {
    const sheet = args.state.workbook.getSheetById(ref.sheetId)
    if (!sheet) {
      throw new Error(`Unknown sheet id: ${ref.sheetId}`)
    }
    const address = formatAddress(ref.mutation.row, ref.mutation.col)
    const existingCellIndex = sheet.grid.get(ref.mutation.row, ref.mutation.col)
    const cellIndex = existingCellIndex === -1 ? undefined : existingCellIndex
    if (cellIndex === undefined) {
      return { kind: 'clearCell', sheetName: sheet.name, address }
    }
    const cellStore = args.state.workbook.cellStore
    const formulaId = cellStore.formulaIds[cellIndex] ?? 0
    if (formulaId === 0) {
      const tag = cellStore.tags[cellIndex]
      if (tag === ValueTag.Number) {
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: cellStore.numbers[cellIndex] ?? 0,
        }
      }
      if (tag === ValueTag.Boolean) {
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: (cellStore.numbers[cellIndex] ?? 0) !== 0,
        }
      }
      if (tag === ValueTag.Empty || tag === ValueTag.Error || tag === undefined) {
        return { kind: 'clearCell', sheetName: sheet.name, address }
      }
    }
    const runtimeFormula = args.state.formulas.get(cellIndex)
    if (runtimeFormula?.source !== undefined) {
      return {
        kind: 'setCellFormula',
        sheetName: sheet.name,
        address,
        formula: runtimeFormula.source,
      }
    }
    const snapshot = args.getCellByIndex(cellIndex)
    if (snapshot.formula !== undefined) {
      return {
        kind: 'setCellFormula',
        sheetName: sheet.name,
        address,
        formula: snapshot.formula,
      }
    }
    switch (snapshot.value.tag) {
      case ValueTag.Empty:
      case ValueTag.Error:
        return (snapshot.flags & CellFlags.AuthoredBlank) !== 0
          ? {
              kind: 'setCellValue',
              sheetName: sheet.name,
              address,
              value: null,
            }
          : { kind: 'clearCell', sheetName: sheet.name, address }
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return {
          kind: 'setCellValue',
          sheetName: sheet.name,
          address,
          value: snapshot.value.value,
        }
    }
  }

  const readStoredCellState = (sheetName: string, address: string): ComparableCellState => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return { value: null, format: null, authoredBlank: false }
    }
    const snapshot = args.getCellByIndex(cellIndex)
    const format = args.state.workbook.getCellFormat(cellIndex) ?? null
    if (snapshot.formula !== undefined) {
      return { formula: snapshot.formula, value: null, format }
    }
    const authoredBlank = ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0
    switch (snapshot.value.tag) {
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return { value: snapshot.value.value, format }
      case ValueTag.Empty:
      case ValueTag.Error:
        return { value: null, format, authoredBlank }
    }
  }

  const readDesiredCellState = (
    targetSheetName: string,
    targetAddress: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
    formatOverride: string | null = snapshot.format ?? null,
  ): ComparableCellState => {
    if (snapshot.formula !== undefined) {
      return {
        formula:
          sourceSheetName && sourceAddress
            ? translateFormulaForTarget(snapshot.formula, sourceSheetName, sourceAddress, targetSheetName, targetAddress)
            : snapshot.formula,
        value: null,
        format: formatOverride,
      }
    }
    switch (snapshot.value.tag) {
      case ValueTag.Number:
      case ValueTag.Boolean:
      case ValueTag.String:
        return { value: snapshot.value.value, format: formatOverride }
      case ValueTag.Empty:
      case ValueTag.Error:
        return {
          value: null,
          format: formatOverride,
          authoredBlank: (snapshot.flags & CellFlags.AuthoredBlank) !== 0,
        }
    }
  }

  const hasStoredCellContent = (sheetName: string, address: string): boolean => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return false
    }
    const snapshot = args.getCellByIndex(cellIndex)
    if (snapshot.formula !== undefined) {
      return true
    }
    return snapshot.value.tag !== ValueTag.Empty || (snapshot.flags & CellFlags.AuthoredBlank) !== 0
  }

  const hasStoredCellState = (sheetName: string, address: string): boolean => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return false
    }
    const snapshot = args.getCellByIndex(cellIndex)
    return (
      snapshot.formula !== undefined ||
      snapshot.value.tag !== ValueTag.Empty ||
      args.state.workbook.getCellFormat(cellIndex) !== undefined ||
      ((args.state.workbook.cellStore.flags[cellIndex] ?? 0) & CellFlags.AuthoredBlank) !== 0
    )
  }

  const shouldApplyCellState = (
    targetSheetName: string,
    targetAddress: string,
    snapshot: CellSnapshot,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): boolean => {
    const current = readStoredCellState(targetSheetName, targetAddress)
    const desired = readDesiredCellState(targetSheetName, targetAddress, snapshot, sourceSheetName, sourceAddress)
    return (
      current.formula !== desired.formula ||
      current.value !== desired.value ||
      current.format !== desired.format ||
      (current.authoredBlank ?? false) !== (desired.authoredBlank ?? false)
    )
  }

  const buildFastMutationHistoryFromRefs = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells: number,
    options: {
      includeUndoOps?: boolean
    } = {},
  ): FastMutationHistoryResult => {
    if (refs.length === 1) {
      const ref = refs[0]!
      const forwardOp = cellMutationRefToEngineOp(args.state.workbook, ref)
      const inverseOp = restoreCellOpFromRef(ref)
      return {
        forward: createLazySingleOpTransactionRecord(forwardOp, potentialNewCells),
        inverse: createLazySingleOpTransactionRecord(inverseOp, 1),
        undoOps: options.includeUndoOps === false ? null : [structuredClone(inverseOp)],
      }
    }
    const forwardOps: EngineOp[] = Array.from({ length: refs.length })
    for (let index = 0; index < refs.length; index += 1) {
      const ref = refs[index]!
      forwardOps[index] = cellMutationRefToEngineOp(args.state.workbook, ref)
    }

    const inverseOps: EngineOp[] = []
    for (let index = refs.length - 1; index >= 0; index -= 1) {
      inverseOps.push(restoreCellOpFromRef(refs[index]!))
    }

    return {
      forward: { kind: 'ops', ops: forwardOps, potentialNewCells },
      inverse: { kind: 'ops', ops: inverseOps, potentialNewCells: refs.length },
      undoOps: options.includeUndoOps === false ? null : structuredClone(inverseOps),
    }
  }

  const inverseOpsFor = (op: EngineOp): EngineOp[] => {
    const captureStructuralWorkbookMetadataOps = (): EngineOp[] => {
      const restoredOps: EngineOp[] = []
      args.state.workbook.listDefinedNames().forEach(({ name, value }) => {
        restoredOps.push({
          kind: 'upsertDefinedName',
          name,
          value: structuredClone(value),
        })
      })
      args.state.workbook.listTables().forEach((table) => {
        restoredOps.push({
          kind: 'upsertTable',
          table: {
            name: table.name,
            sheetName: table.sheetName,
            startAddress: table.startAddress,
            endAddress: table.endAddress,
            columnNames: [...table.columnNames],
            headerRow: table.headerRow,
            totalsRow: table.totalsRow,
          },
        })
      })
      args.state.workbook.listSpills().forEach((spill) => {
        restoredOps.push({
          kind: 'upsertSpillRange',
          sheetName: spill.sheetName,
          address: spill.address,
          rows: spill.rows,
          cols: spill.cols,
        })
      })
      args.state.workbook.listPivots().forEach((pivot) => {
        restoredOps.push({
          kind: 'upsertPivotTable',
          name: pivot.name,
          sheetName: pivot.sheetName,
          address: pivot.address,
          source: { ...pivot.source },
          groupBy: [...pivot.groupBy],
          values: pivot.values.map((value) => Object.assign({}, value)),
          rows: pivot.rows,
          cols: pivot.cols,
        })
      })
      args.state.workbook.listCharts().forEach((chart) => {
        restoredOps.push({
          kind: 'upsertChart',
          chart: structuredClone(chart),
        })
      })
      args.state.workbook.listImages().forEach((image) => {
        restoredOps.push({
          kind: 'upsertImage',
          image: structuredClone(image),
        })
      })
      args.state.workbook.listShapes().forEach((shape) => {
        restoredOps.push({
          kind: 'upsertShape',
          shape: structuredClone(shape),
        })
      })
      return restoredOps
    }

    const clearStructuralSheetMetadataOps = (sheetName: string, transform: ReturnType<typeof structuralTransformForOp>): EngineOp[] => {
      const clearedOps: EngineOp[] = []
      args.state.workbook.listFilters(sheetName).forEach((filter) => {
        const range = rewriteRangeForStructuralTransform(filter.range.startAddress, filter.range.endAddress, transform)
        if (!range) {
          return
        }
        clearedOps.push({
          kind: 'clearFilter',
          sheetName,
          range: {
            ...filter.range,
            startAddress: range.startAddress,
            endAddress: range.endAddress,
          },
        })
      })
      args.state.workbook.listSorts(sheetName).forEach((sort) => {
        const range = rewriteRangeForStructuralTransform(sort.range.startAddress, sort.range.endAddress, transform)
        if (!range) {
          return
        }
        clearedOps.push({
          kind: 'clearSort',
          sheetName,
          range: { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
        })
      })
      args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
        const range = rewriteRangeForStructuralTransform(validation.range.startAddress, validation.range.endAddress, transform)
        if (!range) {
          return
        }
        clearedOps.push({
          kind: 'clearDataValidation',
          sheetName,
          range: {
            ...validation.range,
            startAddress: range.startAddress,
            endAddress: range.endAddress,
          },
        })
      })
      args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
        const address = rewriteAddressForStructuralTransform(thread.address, transform)
        if (!address) {
          return
        }
        clearedOps.push({
          kind: 'deleteCommentThread',
          sheetName,
          address,
        })
      })
      args.state.workbook.listNotes(sheetName).forEach((note) => {
        const address = rewriteAddressForStructuralTransform(note.address, transform)
        if (!address) {
          return
        }
        clearedOps.push({
          kind: 'deleteNote',
          sheetName,
          address,
        })
      })
      return clearedOps
    }

    switch (op.kind) {
      case 'upsertWorkbook':
        return [{ kind: 'upsertWorkbook', name: args.state.workbook.workbookName }]
      case 'setWorkbookMetadata': {
        const existing = args.state.workbook.getWorkbookProperty(op.key)
        return [{ kind: 'setWorkbookMetadata', key: op.key, value: existing?.value ?? null }]
      }
      case 'setCalculationSettings':
        return [
          {
            kind: 'setCalculationSettings',
            settings: args.state.workbook.getCalculationSettings(),
          },
        ]
      case 'setVolatileContext':
        return [{ kind: 'setVolatileContext', context: args.state.workbook.getVolatileContext() }]
      case 'upsertSheet': {
        const existing = args.state.workbook.getSheet(op.name)
        if (!existing) {
          return [{ kind: 'deleteSheet', name: op.name }]
        }
        return [{ kind: 'upsertSheet', name: existing.name, order: existing.order }]
      }
      case 'renameSheet': {
        const existing = args.state.workbook.getSheet(op.newName)
        if (!existing) {
          return []
        }
        return [{ kind: 'renameSheet', oldName: op.newName, newName: op.oldName }]
      }
      case 'deleteSheet': {
        const sheet = args.state.workbook.getSheet(op.name)
        if (!sheet) {
          return []
        }
        const restoredOps: EngineOp[] = [{ kind: 'upsertSheet', name: sheet.name, order: sheet.order }]
        restoredOps.push(...sheetMetadataToOps(args.state.workbook, sheet.name))
        args.state.workbook
          .listTables()
          .filter((table) => table.sheetName === sheet.name)
          .forEach((table) => {
            restoredOps.push({
              kind: 'upsertTable',
              table: {
                name: table.name,
                sheetName: table.sheetName,
                startAddress: table.startAddress,
                endAddress: table.endAddress,
                columnNames: [...table.columnNames],
                headerRow: table.headerRow,
                totalsRow: table.totalsRow,
              },
            })
          })
        args.state.workbook
          .listSpills()
          .filter((spill) => spill.sheetName === sheet.name)
          .forEach((spill) => {
            restoredOps.push({
              kind: 'upsertSpillRange',
              sheetName: spill.sheetName,
              address: spill.address,
              rows: spill.rows,
              cols: spill.cols,
            })
          })
        args.state.workbook
          .listPivots()
          .filter((pivot) => pivot.sheetName === sheet.name)
          .forEach((pivot) => {
            restoredOps.push({
              kind: 'upsertPivotTable',
              name: pivot.name,
              sheetName: pivot.sheetName,
              address: pivot.address,
              source: { ...pivot.source },
              groupBy: [...pivot.groupBy],
              values: pivot.values.map((value) => Object.assign({}, value)),
              rows: pivot.rows,
              cols: pivot.cols,
            })
          })
        args.state.workbook
          .listCharts()
          .filter((chart) => chart.sheetName === sheet.name || chart.source.sheetName === sheet.name)
          .forEach((chart) => {
            restoredOps.push({
              kind: 'upsertChart',
              chart: structuredClone(chart),
            })
          })
        args.state.workbook
          .listImages()
          .filter((image) => image.sheetName === sheet.name)
          .forEach((image) => {
            restoredOps.push({
              kind: 'upsertImage',
              image: structuredClone(image),
            })
          })
        args.state.workbook
          .listShapes()
          .filter((shape) => shape.sheetName === sheet.name)
          .forEach((shape) => {
            restoredOps.push({
              kind: 'upsertShape',
              shape: structuredClone(shape),
            })
          })
        restoredOps.push(...args.captureSheetCellState(sheet.name))
        return restoredOps
      }
      case 'insertRows':
        return [{ kind: 'deleteRows', sheetName: op.sheetName, start: op.start, count: op.count }]
      case 'deleteRows': {
        const entries = args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(op.sheetName, transform),
          {
            kind: 'insertRows',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...sheetMetadataToOps(args.state.workbook, op.sheetName, { includeAxisEntries: false }),
          ...args.captureRowRangeCellState(op.sheetName, op.start, op.count),
          ...captureFormulaCellStateForStructuralUndo(op.sheetName, 'row', op.start, op.count),
          ...captureStructuralWorkbookMetadataOps(),
        ]
      }
      case 'moveRows':
        return [
          {
            kind: 'moveRows',
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ]
      case 'insertColumns':
        return [{ kind: 'deleteColumns', sheetName: op.sheetName, start: op.start, count: op.count }]
      case 'deleteColumns': {
        const entries = args.state.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(op.sheetName, transform),
          {
            kind: 'insertColumns',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            entries,
          },
          ...sheetMetadataToOps(args.state.workbook, op.sheetName, { includeAxisEntries: false }),
          ...args.captureColumnRangeCellState(op.sheetName, op.start, op.count),
          ...captureFormulaCellStateForStructuralUndo(op.sheetName, 'column', op.start, op.count),
          ...captureStructuralWorkbookMetadataOps(),
        ]
      }
      case 'moveColumns':
        return [
          {
            kind: 'moveColumns',
            sheetName: op.sheetName,
            start: op.target,
            count: op.count,
            target: op.start,
          },
        ]
      case 'updateRowMetadata': {
        const existing = args.state.workbook.getRowMetadata(op.sheetName, op.start, op.count)
        return [
          {
            kind: 'updateRowMetadata',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ]
      }
      case 'updateColumnMetadata': {
        const existing = args.state.workbook.getColumnMetadata(op.sheetName, op.start, op.count)
        return [
          {
            kind: 'updateColumnMetadata',
            sheetName: op.sheetName,
            start: op.start,
            count: op.count,
            size: existing?.size ?? null,
            hidden: existing?.hidden ?? null,
          },
        ]
      }
      case 'setFreezePane': {
        const existing = args.state.workbook.getFreezePane(op.sheetName)
        if (!existing) {
          return [{ kind: 'clearFreezePane', sheetName: op.sheetName }]
        }
        return [
          {
            kind: 'setFreezePane',
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'clearFreezePane': {
        const existing = args.state.workbook.getFreezePane(op.sheetName)
        if (!existing) {
          return []
        }
        return [
          {
            kind: 'setFreezePane',
            sheetName: op.sheetName,
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'setFilter': {
        const existing = args.state.workbook.getFilter(op.sheetName, op.range)
        if (!existing) {
          return [{ kind: 'clearFilter', sheetName: op.sheetName, range: { ...op.range } }]
        }
        return [{ kind: 'setFilter', sheetName: op.sheetName, range: { ...existing.range } }]
      }
      case 'clearFilter': {
        const existing = args.state.workbook.getFilter(op.sheetName, op.range)
        if (!existing) {
          return []
        }
        return [{ kind: 'setFilter', sheetName: op.sheetName, range: { ...existing.range } }]
      }
      case 'setSort': {
        const existing = args.state.workbook.getSort(op.sheetName, op.range)
        if (!existing) {
          return [{ kind: 'clearSort', sheetName: op.sheetName, range: { ...op.range } }]
        }
        return [
          {
            kind: 'setSort',
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ]
      }
      case 'clearSort': {
        const existing = args.state.workbook.getSort(op.sheetName, op.range)
        if (!existing) {
          return []
        }
        return [
          {
            kind: 'setSort',
            sheetName: op.sheetName,
            range: { ...existing.range },
            keys: existing.keys.map((key) => Object.assign({}, key)),
          },
        ]
      }
      case 'setDataValidation': {
        const existing = args.state.workbook.getDataValidation(op.validation.range.sheetName, op.validation.range)
        if (!existing) {
          return [
            {
              kind: 'clearDataValidation',
              sheetName: op.validation.range.sheetName,
              range: { ...op.validation.range },
            },
          ]
        }
        return [{ kind: 'setDataValidation', validation: structuredClone(existing) }]
      }
      case 'clearDataValidation': {
        const existing = args.state.workbook.getDataValidation(op.sheetName, op.range)
        if (!existing) {
          return []
        }
        return [{ kind: 'setDataValidation', validation: structuredClone(existing) }]
      }
      case 'setSheetProtection': {
        const existing = args.state.workbook.getSheetProtection(op.protection.sheetName)
        if (!existing) {
          return [{ kind: 'clearSheetProtection', sheetName: op.protection.sheetName }]
        }
        return [{ kind: 'setSheetProtection', protection: structuredClone(existing) }]
      }
      case 'clearSheetProtection': {
        const existing = args.state.workbook.getSheetProtection(op.sheetName)
        if (!existing) {
          return []
        }
        return [{ kind: 'setSheetProtection', protection: structuredClone(existing) }]
      }
      case 'upsertConditionalFormat': {
        const existing = args.state.workbook.getConditionalFormat(op.format.id)
        if (!existing) {
          return [
            {
              kind: 'deleteConditionalFormat',
              id: op.format.id,
              sheetName: op.format.range.sheetName,
            },
          ]
        }
        return [{ kind: 'upsertConditionalFormat', format: structuredClone(existing) }]
      }
      case 'deleteConditionalFormat': {
        const existing = args.state.workbook.getConditionalFormat(op.id)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertConditionalFormat', format: structuredClone(existing) }]
      }
      case 'upsertRangeProtection': {
        const existing = args.state.workbook.getRangeProtection(op.protection.id)
        if (!existing) {
          return [
            {
              kind: 'deleteRangeProtection',
              id: op.protection.id,
              sheetName: op.protection.range.sheetName,
            },
          ]
        }
        return [{ kind: 'upsertRangeProtection', protection: structuredClone(existing) }]
      }
      case 'deleteRangeProtection': {
        const existing = args.state.workbook.getRangeProtection(op.id)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertRangeProtection', protection: structuredClone(existing) }]
      }
      case 'upsertCommentThread': {
        const existing = args.state.workbook.getCommentThread(op.thread.sheetName, op.thread.address)
        if (!existing) {
          return [
            {
              kind: 'deleteCommentThread',
              sheetName: op.thread.sheetName,
              address: op.thread.address,
            },
          ]
        }
        return [{ kind: 'upsertCommentThread', thread: structuredClone(existing) }]
      }
      case 'deleteCommentThread': {
        const existing = args.state.workbook.getCommentThread(op.sheetName, op.address)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertCommentThread', thread: structuredClone(existing) }]
      }
      case 'upsertNote': {
        const existing = args.state.workbook.getNote(op.note.sheetName, op.note.address)
        if (!existing) {
          return [
            {
              kind: 'deleteNote',
              sheetName: op.note.sheetName,
              address: op.note.address,
            },
          ]
        }
        return [{ kind: 'upsertNote', note: structuredClone(existing) }]
      }
      case 'deleteNote': {
        const existing = args.state.workbook.getNote(op.sheetName, op.address)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertNote', note: structuredClone(existing) }]
      }
      case 'setCellValue':
      case 'setCellFormula':
      case 'clearCell':
        return args.restoreCellOps(op.sheetName, op.address)
      case 'setCellFormat': {
        const cellIndex = args.state.workbook.getCellIndex(op.sheetName, op.address)
        return [
          {
            kind: 'setCellFormat',
            sheetName: op.sheetName,
            address: op.address,
            format: cellIndex === undefined ? null : (args.state.workbook.getCellFormat(cellIndex) ?? null),
          },
        ]
      }
      case 'upsertCellStyle': {
        const existing = args.state.workbook.getCellStyle(op.style.id)
        if (!existing || existing.id !== op.style.id) {
          return []
        }
        return [{ kind: 'upsertCellStyle', style: cloneCellStyleRecord(existing) }]
      }
      case 'upsertCellNumberFormat': {
        const existing = args.state.workbook.getCellNumberFormat(op.format.id)
        if (!existing || existing.id !== op.format.id) {
          return []
        }
        return [{ kind: 'upsertCellNumberFormat', format: { ...existing } }]
      }
      case 'setStyleRange':
        return restoreStyleRangeOps(args.state.workbook, op.range)
      case 'setFormatRange':
        return restoreFormatRangeOps(args.state.workbook, op.range)
      case 'upsertDefinedName': {
        const existing = args.state.workbook.getDefinedName(op.name)
        if (!existing) {
          return [{ kind: 'deleteDefinedName', name: op.name }]
        }
        return [{ kind: 'upsertDefinedName', name: existing.name, value: existing.value }]
      }
      case 'deleteDefinedName': {
        const existing = args.state.workbook.getDefinedName(op.name)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertDefinedName', name: existing.name, value: existing.value }]
      }
      case 'upsertTable': {
        const existing = args.state.workbook.getTable(op.table.name)
        if (!existing) {
          return [{ kind: 'deleteTable', name: op.table.name }]
        }
        return [
          {
            kind: 'upsertTable',
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ]
      }
      case 'deleteTable': {
        const existing = args.state.workbook.getTable(op.name)
        if (!existing) {
          return []
        }
        return [
          {
            kind: 'upsertTable',
            table: {
              name: existing.name,
              sheetName: existing.sheetName,
              startAddress: existing.startAddress,
              endAddress: existing.endAddress,
              columnNames: [...existing.columnNames],
              headerRow: existing.headerRow,
              totalsRow: existing.totalsRow,
            },
          },
        ]
      }
      case 'upsertSpillRange': {
        const existing = args.state.workbook.getSpill(op.sheetName, op.address)
        if (!existing) {
          return [{ kind: 'deleteSpillRange', sheetName: op.sheetName, address: op.address }]
        }
        return [
          {
            kind: 'upsertSpillRange',
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'deleteSpillRange': {
        const existing = args.state.workbook.getSpill(op.sheetName, op.address)
        if (!existing) {
          return []
        }
        return [
          {
            kind: 'upsertSpillRange',
            sheetName: existing.sheetName,
            address: existing.address,
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'upsertPivotTable': {
        const existing = args.state.workbook.getPivot(op.sheetName, op.address)
        if (!existing) {
          return [{ kind: 'deletePivotTable', sheetName: op.sheetName, address: op.address }]
        }
        return [
          {
            kind: 'upsertPivotTable',
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((v) => Object.assign({}, v)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'deletePivotTable': {
        const existing = args.state.workbook.getPivot(op.sheetName, op.address)
        if (!existing) {
          return []
        }
        return [
          {
            kind: 'upsertPivotTable',
            name: existing.name,
            sheetName: existing.sheetName,
            address: existing.address,
            source: { ...existing.source },
            groupBy: [...existing.groupBy],
            values: existing.values.map((value) => Object.assign({}, value)),
            rows: existing.rows,
            cols: existing.cols,
          },
        ]
      }
      case 'upsertChart': {
        const existing = args.state.workbook.getChart(op.chart.id)
        if (!existing) {
          return [{ kind: 'deleteChart', id: op.chart.id }]
        }
        return [{ kind: 'upsertChart', chart: structuredClone(existing) }]
      }
      case 'deleteChart': {
        const existing = args.state.workbook.getChart(op.id)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertChart', chart: structuredClone(existing) }]
      }
      case 'upsertImage': {
        const existing = args.state.workbook.getImage(op.image.id)
        if (!existing) {
          return [{ kind: 'deleteImage', id: op.image.id }]
        }
        return [{ kind: 'upsertImage', image: structuredClone(existing) }]
      }
      case 'deleteImage': {
        const existing = args.state.workbook.getImage(op.id)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertImage', image: structuredClone(existing) }]
      }
      case 'upsertShape': {
        const existing = args.state.workbook.getShape(op.shape.id)
        if (!existing) {
          return [{ kind: 'deleteShape', id: op.shape.id }]
        }
        return [{ kind: 'upsertShape', shape: structuredClone(existing) }]
      }
      case 'deleteShape': {
        const existing = args.state.workbook.getShape(op.id)
        if (!existing) {
          return []
        }
        return [{ kind: 'upsertShape', shape: structuredClone(existing) }]
      }
      default: {
        const exhaustive: never = op
        return exhaustive
      }
    }
  }

  const buildInverseOps = (ops: readonly EngineOp[]): EngineOp[] => {
    const inverseOps: EngineOp[] = []
    for (let index = ops.length - 1; index >= 0; index -= 1) {
      const op = ops[index]
      if (op !== undefined) {
        inverseOps.push(...inverseOpsFor(op))
      }
    }
    return inverseOps
  }

  const captureFormulaCellStateForStructuralUndo = (
    sheetName: string,
    axis: 'row' | 'column',
    start: number,
    count: number,
  ): EngineOp[] => {
    const captured: EngineOp[] = []
    args.state.formulas.forEach((_formula, cellIndex) => {
      const ownerSheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
      if (ownerSheetId === undefined) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(ownerSheetId)
      if (!ownerSheetName) {
        return
      }
      const axisIndex = axis === 'row' ? args.state.workbook.cellStore.rows[cellIndex] : args.state.workbook.cellStore.cols[cellIndex]
      if (ownerSheetName === sheetName && axisIndex !== undefined && axisIndex >= start && axisIndex < start + count) {
        return
      }
      captured.push(...args.captureStoredCellOps(cellIndex, ownerSheetName, args.state.workbook.getAddress(cellIndex)))
    })
    return captured
  }

  const canonicalizeForwardOps = (ops: readonly EngineOp[]): EngineOp[] =>
    ops.map((op) => {
      if (op.kind === 'insertRows') {
        return op.entries
          ? { ...op, entries: op.entries.map((entry) => ({ ...entry })) }
          : {
              ...op,
              entries: args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count),
            }
      }

      if (op.kind === 'insertColumns') {
        return op.entries
          ? { ...op, entries: op.entries.map((entry) => ({ ...entry })) }
          : {
              ...op,
              entries: args.state.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count),
            }
      }

      return structuredClone(op)
    })

  const executeTransactionNow = (record: TransactionRecord, source: 'local' | 'restore' | 'undo' | 'redo'): void => {
    if (record.kind === 'ops' && record.ops.length === 0) {
      return
    }
    const batch = createBatch(args.state.replicaState, record.kind === 'single-op' ? [record.op] : record.ops)
    args.applyBatchNow(batch, source, record.potentialNewCells)
  }

  const executeLocalNowWithCustomApply = (
    ops: EngineOp[],
    potentialNewCells: number | undefined,
    applyForward: (forward: TransactionRecord) => void,
    options: {
      returnUndoOps?: boolean
      reuseForwardOps?: boolean
      preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[]
    } = {},
  ): readonly EngineOp[] | null => {
    if (ops.length === 0) {
      return null
    }
    const forward: TransactionRecord =
      potentialNewCells === undefined
        ? options.preparedCellAddressesByOpIndex
          ? {
              kind: 'ops',
              ops,
              preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
            }
          : { kind: 'ops', ops }
        : options.preparedCellAddressesByOpIndex
          ? {
              kind: 'ops',
              ops,
              potentialNewCells,
              preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
            }
          : { kind: 'ops', ops, potentialNewCells }
    const baseFastHistoryArgs: Parameters<typeof tryBuildFastMutationHistory>[0] =
      potentialNewCells === undefined
        ? {
            workbook: args.state.workbook,
            getCellByIndex: args.getCellByIndex,
            ops,
            cloneForwardOps: options.reuseForwardOps !== true,
          }
        : {
            workbook: args.state.workbook,
            getCellByIndex: args.getCellByIndex,
            ops,
            potentialNewCells,
            includeUndoOps: options.returnUndoOps !== false,
            cloneForwardOps: options.reuseForwardOps !== true,
          }
    const fastHistoryArgs: Parameters<typeof tryBuildFastMutationHistory>[0] = options.preparedCellAddressesByOpIndex
      ? {
          ...baseFastHistoryArgs,
          preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
        }
      : baseFastHistoryArgs
    const fastHistory = tryBuildFastMutationHistory(fastHistoryArgs)
    const inverse: TransactionRecord = fastHistory?.inverse ?? {
      kind: 'ops',
      ops: buildInverseOps(ops),
      potentialNewCells: ops.length,
    }
    applyForward(forward)
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward:
          fastHistory?.forward ??
          (potentialNewCells === undefined
            ? options.preparedCellAddressesByOpIndex
              ? {
                  kind: 'ops',
                  ops: canonicalizeForwardOps(ops),
                  preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
                }
              : { kind: 'ops', ops: canonicalizeForwardOps(ops) }
            : options.preparedCellAddressesByOpIndex
              ? {
                  kind: 'ops',
                  ops: canonicalizeForwardOps(ops),
                  potentialNewCells,
                  preparedCellAddressesByOpIndex: options.preparedCellAddressesByOpIndex,
                }
              : { kind: 'ops', ops: canonicalizeForwardOps(ops), potentialNewCells }),
        inverse,
      })
      args.state.redoStack.length = 0
    }
    if (options.returnUndoOps === false) {
      return null
    }
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(inverse)
  }

  const executeLocalCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    potentialNewCells?: number,
    options: {
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null => {
    if (refs.length === 0) {
      return null
    }
    const nextRefs = options.reuseRefs ? refs : refs.map((ref) => cloneCellMutationRef(ref))
    const nextPotentialNewCells = potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs)
    if (nextRefs.length === 1 && options.returnUndoOps === false) {
      const ref = nextRefs[0]!
      const forwardOp = cellMutationRefToEngineOp(args.state.workbook, ref)
      const inverseOp = restoreCellOpFromRef(ref)
      const batch = shouldCreateLocalBatch() ? createBatch(args.state.replicaState, [forwardOp]) : null
      args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord(forwardOp, nextPotentialNewCells),
          inverse: createLazySingleOpTransactionRecord(inverseOp, 1),
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    if (!shouldCreateLocalBatch() && nextRefs.length > 1 && options.returnUndoOps === false) {
      const inverse: TransactionRecord = {
        kind: 'ops',
        ops: nextRefs.toReversed().map((ref) => restoreCellOpFromRef(ref)),
        potentialNewCells: nextRefs.length,
      }
      args.applyCellMutationsAtBatchNow(nextRefs, null, 'local', nextPotentialNewCells)
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazyCellMutationTransactionRecord(args.state.workbook, nextRefs, nextPotentialNewCells),
          inverse,
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    const fastHistory = buildFastMutationHistoryFromRefs(nextRefs, nextPotentialNewCells, {
      includeUndoOps: options.returnUndoOps !== false,
    })
    const inverse: TransactionRecord = fastHistory?.inverse ?? {
      kind: 'ops',
      ops: buildInverseOps(transactionRecordOps(fastHistory.forward)),
      potentialNewCells: transactionRecordOps(fastHistory.forward).length,
    }
    const batch = shouldCreateLocalBatch() ? createBatch(args.state.replicaState, [...transactionRecordOps(fastHistory.forward)]) : null
    args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: fastHistory?.forward ?? {
          kind: 'ops',
          ops: canonicalizeForwardOps([...transactionRecordOps(fastHistory.forward)]),
          potentialNewCells: nextPotentialNewCells,
        },
        inverse,
      })
      args.state.redoStack.length = 0
    }
    if (options.returnUndoOps === false) {
      return null
    }
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(inverse)
  }

  const applyCellMutationsAtNow = (
    refs: readonly EngineCellMutationRef[],
    options: {
      captureUndo?: boolean
      potentialNewCells?: number
      source?: 'local' | 'restore'
      returnUndoOps?: boolean
      reuseRefs?: boolean
    } = {},
  ): readonly EngineOp[] | null => {
    const source = options.source ?? 'restore'
    const captureUndo = options.captureUndo ?? source === 'local'
    if (captureUndo) {
      const executeOptions: {
        returnUndoOps?: boolean
        reuseRefs?: boolean
      } = {}
      if (options.returnUndoOps !== undefined) {
        executeOptions.returnUndoOps = options.returnUndoOps
      }
      if (options.reuseRefs !== undefined) {
        executeOptions.reuseRefs = options.reuseRefs
      }
      return executeLocalCellMutationsAtNow(refs, options.potentialNewCells, {
        ...executeOptions,
      })
    }
    if (refs.length === 0) {
      return null
    }
    const nextRefs = options.reuseRefs ? refs : refs.map((ref) => cloneCellMutationRef(ref))
    const nextPotentialNewCells = options.potentialNewCells ?? countPotentialNewCellsForMutationRefs(nextRefs)
    const forwardOps = source === 'restore' ? emptyBatchOps : nextRefs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref))
    const batch =
      source === 'local' && shouldCreateLocalBatch()
        ? createBatch(args.state.replicaState, forwardOps)
        : source === 'restore'
          ? null
          : createBatch(args.state.replicaState, forwardOps)
    args.applyCellMutationsAtBatchNow(nextRefs, batch, source, nextPotentialNewCells)
    return null
  }

  return {
    executeTransactionNow: executeTransactionNow,
    executeTransaction(record, source) {
      return Effect.try({
        try: () => {
          executeTransactionNow(record, source)
        },
        catch: (cause) =>
          new EngineMutationError({
            message: `Failed to execute ${source} transaction`,
            cause,
          }),
      })
    },
    executeLocalNow(ops, potentialNewCells) {
      return executeLocalNowWithCustomApply(
        ops,
        potentialNewCells,
        (forward) => {
          executeTransactionNow(forward, 'local')
        },
        { returnUndoOps: true, reuseForwardOps: false },
      )
    },
    executeLocalCellMutationsAtNow(refs, potentialNewCells) {
      return executeLocalCellMutationsAtNow(refs, potentialNewCells)
    },
    applyCellMutationsAtNow(refs, options = {}) {
      return applyCellMutationsAtNow(refs, options)
    },
    applyCellMutationsAt(refs, options = {}) {
      return Effect.try({
        try: () => applyCellMutationsAtNow(refs, options),
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to apply cell mutations',
            cause,
          }),
      })
    },
    executeLocal(ops, potentialNewCells) {
      return Effect.try({
        try: () => this.executeLocalNow(ops, potentialNewCells),
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to execute local transaction',
            cause,
          }),
      })
    },
    applyOpsNow(ops, options = {}) {
      const nextOps = options.trusted ? Array.from(ops) : structuredClone([...ops])
      if (nextOps.length === 0) {
        return null
      }
      if (options.captureUndo) {
        return this.executeLocalNow(nextOps, options.potentialNewCells)
      }
      executeTransactionNow(
        options.potentialNewCells === undefined
          ? { kind: 'ops', ops: nextOps }
          : { kind: 'ops', ops: nextOps, potentialNewCells: options.potentialNewCells },
        options.source ?? 'restore',
      )
      return null
    },
    applyOps(ops, options = {}) {
      return Effect.try({
        try: () => this.applyOpsNow(ops, options),
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to apply engine operations',
            cause,
          }),
      })
    },
    captureUndoOps(mutate) {
      return Effect.try({
        try: () => {
          const previousUndoDepth = args.state.undoStack.length
          const result = mutate()
          if (args.state.undoStack.length === previousUndoDepth) {
            return {
              result,
              undoOps: null,
            }
          }
          if (args.state.undoStack.length === previousUndoDepth + 1) {
            const inverse = args.state.undoStack.at(-1)!.inverse
            return {
              result,
              undoOps: cloneTransactionRecordOps(inverse),
            }
          }
          throw new Error('Expected a single local transaction while capturing undo ops')
        },
        catch: (cause) =>
          new EngineMutationError({
            message: 'Failed to capture undo ops',
            cause,
          }),
      })
    },
    setRangeValues(range, values) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const expectedHeight = bounds.endRow - bounds.startRow + 1
          const expectedWidth = bounds.endCol - bounds.startCol + 1
          if (values.length !== expectedHeight || values.some((row) => row.length !== expectedWidth)) {
            throw new Error('setRangeValues requires a value matrix that exactly matches the target range')
          }

          const ops: EngineOp[] = []
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              const address = formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset)
              const current = readStoredCellState(range.sheetName, address)
              const nextValue = values[rowOffset]![colOffset] ?? null
              if (current.formula === undefined && current.value === nextValue) {
                continue
              }
              ops.push({
                kind: 'setCellValue',
                sheetName: range.sheetName,
                address,
                value: nextValue,
              })
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(this.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to set range values', cause),
            cause,
          }),
      })
    },
    setRangeFormulas(range, formulas) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const expectedHeight = bounds.endRow - bounds.startRow + 1
          const expectedWidth = bounds.endCol - bounds.startCol + 1
          if (formulas.length !== expectedHeight || formulas.some((row) => row.length !== expectedWidth)) {
            throw new Error('setRangeFormulas requires a formula matrix that exactly matches the target range')
          }

          const opCount = expectedHeight * expectedWidth
          const ops = Array.from<EngineOp>({ length: opCount })
          let opIndex = 0
          for (let rowOffset = 0; rowOffset < expectedHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < expectedWidth; colOffset += 1) {
              ops[opIndex] = {
                kind: 'setCellFormula',
                sheetName: range.sheetName,
                address: formatAddress(bounds.startRow + rowOffset, bounds.startCol + colOffset),
                formula: formulas[rowOffset]![colOffset] ?? '',
              }
              opIndex += 1
            }
          }
          Effect.runSync(this.executeLocal(ops, opCount))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to set range formulas', cause),
            cause,
          }),
      })
    },
    clearRange(range) {
      return Effect.try({
        try: () => {
          const bounds = normalizeRange(range)
          const ops: EngineOp[] = []
          for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
            for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
              const address = formatAddress(row, col)
              if (!hasStoredCellContent(range.sheetName, address)) {
                continue
              }
              ops.push({
                kind: 'clearCell',
                sheetName: range.sheetName,
                address,
              })
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(this.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to clear range', cause),
            cause,
          }),
      })
    },
    fillRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceMatrix.length
          const sourceWidth = sourceMatrix[0]?.length ?? 0
          if (sourceHeight === 0 || sourceWidth === 0) {
            return
          }

          const ops: EngineOp[] = []
          for (let row = targetBounds.startRow; row <= targetBounds.endRow; row += 1) {
            for (let col = targetBounds.startCol; col <= targetBounds.endCol; col += 1) {
              const sourceRowOffset = (row - targetBounds.startRow) % sourceHeight
              const sourceColOffset = (col - targetBounds.startCol) % sourceWidth
              const sourceCell = getMatrixCell(sourceMatrix, sourceRowOffset, sourceColOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + sourceRowOffset, sourceBounds.startCol + sourceColOffset)
              const nextAddress = formatAddress(row, col)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(this.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to fill range', cause),
            cause,
          }),
      })
    },
    copyRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error('copyRange requires source and target dimensions to match exactly')
          }

          const ops: EngineOp[] = []
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
              const sourceCell = getMatrixCell(sourceMatrix, rowOffset, colOffset)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(this.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to copy range', cause),
            cause,
          }),
      })
    },
    moveRange(source, target) {
      return Effect.try({
        try: () => {
          const sourceMatrix = args.readRangeCells(source)
          const targetBounds = normalizeRange(target)
          const sourceBounds = normalizeRange(source)
          const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
          const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
          const targetHeight = targetBounds.endRow - targetBounds.startRow + 1
          const targetWidth = targetBounds.endCol - targetBounds.startCol + 1
          if (sourceHeight !== targetHeight || sourceWidth !== targetWidth) {
            throw new Error('moveRange requires source and target dimensions to match exactly')
          }
          if (
            source.sheetName === target.sheetName &&
            sourceBounds.startRow === targetBounds.startRow &&
            sourceBounds.endRow === targetBounds.endRow &&
            sourceBounds.startCol === targetBounds.startCol &&
            sourceBounds.endCol === targetBounds.endCol
          ) {
            return
          }

          const ops: EngineOp[] = []
          for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
            for (let col = sourceBounds.startCol; col <= sourceBounds.endCol; col += 1) {
              const address = formatAddress(row, col)
              if (!hasStoredCellState(source.sheetName, address)) {
                continue
              }
              ops.push({
                kind: 'clearCell',
                sheetName: source.sheetName,
                address,
              })
            }
          }
          for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
            for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
              const nextAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
              const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
              const sourceCell = getMatrixCell(sourceMatrix, rowOffset, colOffset)
              if (!shouldApplyCellState(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress)) {
                continue
              }
              ops.push(...args.toCellStateOps(target.sheetName, nextAddress, sourceCell, source.sheetName, sourceAddress))
            }
          }
          if (ops.length === 0) {
            return
          }
          Effect.runSync(this.executeLocal(ops, ops.length))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to move range', cause),
            cause,
          }),
      })
    },
    importSheetCsv(sheetName, csv) {
      return Effect.try({
        try: () => {
          const rows = parseCsv(csv)
          const existingSheet = args.state.workbook.getSheet(sheetName)
          const order = existingSheet?.order ?? args.state.workbook.sheetsByName.size
          const ops: EngineOp[] = []
          let potentialNewCells = 0

          if (existingSheet) {
            ops.push({ kind: 'deleteSheet', name: sheetName })
          }
          ops.push({ kind: 'upsertSheet', name: sheetName, order })

          rows.forEach((row, rowIndex) => {
            row.forEach((raw, colIndex) => {
              const parsed = parseCsvCellInput(raw)
              if (!parsed) {
                return
              }
              const address = formatAddress(rowIndex, colIndex)
              if (parsed.formula !== undefined) {
                ops.push({ kind: 'setCellFormula', sheetName, address, formula: parsed.formula })
                potentialNewCells += 1
                return
              }
              ops.push({ kind: 'setCellValue', sheetName, address, value: parsed.value ?? null })
              potentialNewCells += 1
            })
          })

          Effect.runSync(this.executeLocal(ops, potentialNewCells))
        },
        catch: (cause) =>
          new EngineMutationError({
            message: mutationErrorMessage('Failed to import sheet CSV', cause),
            cause,
          }),
      })
    },
    renderCommit(ops) {
      return Effect.flatMap(
        Effect.try({
          try: () => {
            const maxEngineOpCount = ops.length * 2
            const engineOps: EngineOp[] = []
            engineOps.length = maxEngineOpCount
            const preparedCellAddressesByOpIndex: Array<PreparedCellAddress | null> = []
            preparedCellAddressesByOpIndex.length = maxEngineOpCount
            let engineOpCount = 0
            let potentialNewCells = 0
            const pushEngineOp = (engineOp: EngineOp, preparedCellAddress: PreparedCellAddress | null = null): void => {
              engineOps[engineOpCount] = engineOp
              preparedCellAddressesByOpIndex[engineOpCount] = preparedCellAddress
              engineOpCount += 1
            }
            for (let index = 0; index < ops.length; index += 1) {
              const op = ops[index]
              if (!op) {
                continue
              }
              switch (op.kind) {
                case 'upsertWorkbook':
                  if (op.name) {
                    pushEngineOp({ kind: 'upsertWorkbook', name: op.name })
                  }
                  break
                case 'upsertSheet':
                  if (op.name) {
                    pushEngineOp({ kind: 'upsertSheet', name: op.name, order: op.order ?? 0 })
                  }
                  break
                case 'renameSheet':
                  if (op.oldName && op.newName) {
                    pushEngineOp({
                      kind: 'renameSheet',
                      oldName: op.oldName,
                      newName: op.newName,
                    })
                  }
                  break
                case 'deleteSheet':
                  if (op.name) {
                    pushEngineOp({ kind: 'deleteSheet', name: op.name })
                  }
                  break
                case 'upsertCell': {
                  if (!op.sheetName || !op.addr) {
                    break
                  }
                  const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
                  if (op.formula !== undefined) {
                    pushEngineOp(
                      {
                        kind: 'setCellFormula',
                        sheetName: op.sheetName,
                        address: op.addr,
                        formula: op.formula,
                      },
                      { row: preparedCellAddress.row, col: preparedCellAddress.col },
                    )
                  } else {
                    pushEngineOp(
                      {
                        kind: 'setCellValue',
                        sheetName: op.sheetName,
                        address: op.addr,
                        value: op.value ?? null,
                      },
                      { row: preparedCellAddress.row, col: preparedCellAddress.col },
                    )
                  }
                  potentialNewCells += 1
                  if (op.format !== undefined) {
                    pushEngineOp({
                      kind: 'setCellFormat',
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: op.format,
                    })
                  }
                  break
                }
                case 'deleteCell': {
                  if (op.sheetName && op.addr) {
                    const preparedCellAddress = parseCellAddress(op.addr, op.sheetName)
                    pushEngineOp(
                      {
                        kind: 'clearCell',
                        sheetName: op.sheetName,
                        address: op.addr,
                      },
                      { row: preparedCellAddress.row, col: preparedCellAddress.col },
                    )
                    pushEngineOp({
                      kind: 'setCellFormat',
                      sheetName: op.sheetName,
                      address: op.addr,
                      format: null,
                    })
                  }
                  break
                }
              }
            }
            engineOps.length = engineOpCount
            preparedCellAddressesByOpIndex.length = engineOpCount
            return { engineOps, potentialNewCells, preparedCellAddressesByOpIndex }
          },
          catch: (cause) =>
            new EngineMutationError({
              message: 'Failed to normalize render commit operations',
              cause,
            }),
        }),
        ({ engineOps, potentialNewCells, preparedCellAddressesByOpIndex }) =>
          Effect.try({
            try: () => {
              executeLocalNowWithCustomApply(
                engineOps,
                potentialNewCells,
                (forward) => {
                  const batchOps = forward.kind === 'single-op' ? [forward.op] : forward.ops
                  args.applyBatchNow(
                    createBatch(args.state.replicaState, batchOps),
                    'local',
                    forward.potentialNewCells,
                    preparedCellAddressesByOpIndex,
                  )
                },
                {
                  returnUndoOps: false,
                  reuseForwardOps: true,
                  preparedCellAddressesByOpIndex,
                },
              )
            },
            catch: (cause) =>
              new EngineMutationError({
                message: 'Failed to execute render commit transaction',
                cause,
              }),
          }),
      ).pipe(Effect.asVoid)
    },
  }
}
