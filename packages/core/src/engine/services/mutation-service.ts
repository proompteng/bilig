import { Effect } from 'effect'
import type { EngineOp, EngineOpBatch } from '@bilig/workbook-domain'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import { structuralTransformForOp } from '../../engine-structural-utils.js'
import { sheetMetadataToOps } from '../../engine-snapshot-utils.js'
import { createBatch } from '../../replica-state.js'
import type { WorkbookStore } from '../../workbook-store.js'
import {
  cellMutationRefToEngineOp,
  cloneCellMutationRef,
  countPotentialNewCellsForMutationRefs,
  type EngineCellMutationRef,
  type EngineExistingLiteralCellMutationRef,
  type EngineExistingNumericCellMutationRef,
  type EngineExistingNumericCellMutationResult,
} from '../../cell-mutations-at.js'
import type {
  EngineRuntimeState,
  PreparedCellAddress,
  RuntimeStructuralFormulaSourceTransform,
  TransactionRecord,
} from '../runtime-state.js'
import { EngineMutationError } from '../errors.js'
import { tryBuildFastMutationHistory, type FastMutationHistoryResult } from './mutation-history-fast-path.js'
import {
  cloneTransactionRecordOps,
  createLazyCellMutationTransactionRecord,
  createLazySingleOpTransactionRecord,
  createSingleExistingNumericCellMutationTransactionRecord,
  singleExistingNumericCellMutationRecordToRef,
  transactionRecordOps,
} from './mutation-transaction-records.js'
import { normalizeRenderCommitOps } from './mutation-render-commit-normalizer.js'
import { inverseMutationStructuralInsertOp, isMutationStructuralInsertOp } from './mutation-cell-content-helpers.js'
import { createMutationCellRestoreHistoryHelpers, tryMutationCellRefsFromOps } from './mutation-cell-restore-history.js'
import { captureStructuralWorkbookMetadataOps, clearStructuralSheetMetadataOps } from './mutation-structural-metadata-ops.js'
import { buildMutationMetadataInverseOps } from './mutation-inverse-metadata-ops.js'
import { createMutationStructuralDeleteInverseHelpers } from './mutation-structural-delete-inverse.js'
import type { EngineMutationService } from './mutation-service-types.js'
import { tryExecuteMutationRenderCommitFastPath } from './mutation-render-commit-fast-path.js'
import { createMutationRangeOperations } from './mutation-range-operations.js'

export type { EngineMutationService } from './mutation-service-types.js'

type MutationServiceCapturedInverseKind = 'deleteSheet' | 'deleteRows' | 'deleteColumns' | 'setCellValue' | 'setCellFormula' | 'clearCell'

type MutationServiceCapturedInverseOp = Extract<EngineOp, { kind: MutationServiceCapturedInverseKind }>

const mutationServiceCapturedInverseKinds: ReadonlySet<EngineOp['kind']> = new Set([
  'deleteSheet',
  'deleteRows',
  'deleteColumns',
  'setCellValue',
  'setCellFormula',
  'clearCell',
])

function isMutationServiceCapturedInverseOp(op: EngineOp): op is MutationServiceCapturedInverseOp {
  return mutationServiceCapturedInverseKinds.has(op.kind)
}

export function createEngineMutationService(args: {
  readonly state: Pick<
    EngineRuntimeState,
    | 'replicaState'
    | 'batchListeners'
    | 'formulas'
    | 'undoStack'
    | 'redoStack'
    | 'counters'
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
  readonly getFormulaFamilyStructuralSourceTransform?: (cellIndex: number) => RuntimeStructuralFormulaSourceTransform | undefined
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
  readonly applyExistingNumericCellMutationAtNow?: (
    request: EngineExistingNumericCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly applyExistingLiteralCellMutationAtNow?: (
    request: EngineExistingLiteralCellMutationRef,
  ) => EngineExistingNumericCellMutationResult | null
  readonly hasExternallyVisibleLocalMutationObservers?: () => boolean
}): EngineMutationService {
  const emptyBatchOps: EngineOp[] = []
  const shouldCreateLocalBatch = (): boolean =>
    args.state.trackReplicaVersions ||
    (args.state.batchListeners?.size ?? 0) > 0 ||
    (args.state.getSyncClientConnection?.() ?? null) !== null
  const hasExternallyVisibleBatchRequirement = (): boolean =>
    args.hasExternallyVisibleLocalMutationObservers?.() ??
    ((args.state.batchListeners?.size ?? 0) > 0 || (args.state.getSyncClientConnection?.() ?? null) !== null)

  const {
    restoreCellOpFromRef,
    tryRestoreSimpleCellOpFromStore,
    createLazyInverseCellMutationRecord,
    tryCreateSingleExistingNumericInverseCellMutationRecord,
    buildFastMutationHistoryFromRefs,
  } = createMutationCellRestoreHistoryHelpers({
    workbook: args.state.workbook,
    formulas: args.state.formulas,
    getCellByIndex: args.getCellByIndex,
    ...(args.getFormulaFamilyStructuralSourceTransform
      ? { getFormulaFamilyStructuralSourceTransform: args.getFormulaFamilyStructuralSourceTransform }
      : {}),
  })
  const { captureFormulaCellStateForStructuralUndo, buildStructuralDeleteInverseRecord } = createMutationStructuralDeleteInverseHelpers({
    state: args.state,
    getCellByIndex: args.getCellByIndex,
    toCellStateOps: (sheetName, address, snapshot) => args.toCellStateOps(sheetName, address, snapshot),
    ...(args.getFormulaFamilyStructuralSourceTransform
      ? { getFormulaFamilyStructuralSourceTransform: args.getFormulaFamilyStructuralSourceTransform }
      : {}),
  })

  const tryBuildSingleCellOpHistoryWithoutSnapshot = (
    ops: readonly EngineOp[],
    potentialNewCells: number | undefined,
    includeUndoOps: boolean,
    cloneForwardOp: boolean,
  ): FastMutationHistoryResult | null => {
    if (ops.length !== 1) {
      return null
    }
    const op = ops[0]!
    if (op.kind !== 'setCellValue' && op.kind !== 'setCellFormula' && op.kind !== 'clearCell') {
      return null
    }
    const inverseOp = tryRestoreSimpleCellOpFromStore(op.sheetName, op.address)
    if (inverseOp === null) {
      return null
    }
    const forwardOp = cloneForwardOp ? structuredClone(op) : op
    return {
      forward: createLazySingleOpTransactionRecord(forwardOp, potentialNewCells),
      inverse: createLazySingleOpTransactionRecord(inverseOp, 1),
      undoOps: includeUndoOps ? [structuredClone(inverseOp)] : null,
    }
  }

  const tryCellMutationRefsFromOps = (ops: readonly EngineOp[]): EngineCellMutationRef[] | null =>
    tryMutationCellRefsFromOps(args.state.workbook, ops)

  const inverseOpsFor = (op: EngineOp): EngineOp[] => {
    const metadataInverseOps = buildMutationMetadataInverseOps(args.state.workbook, op)
    if (metadataInverseOps !== undefined) {
      return metadataInverseOps
    }
    if (!isMutationServiceCapturedInverseOp(op)) {
      throw new Error(`Unhandled inverse operation: ${op.kind}`)
    }

    switch (op.kind) {
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
              table: structuredClone(table),
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
      case 'deleteRows': {
        const entries = args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(args.state.workbook, op.sheetName, transform),
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
          ...captureStructuralWorkbookMetadataOps(args.state.workbook),
        ]
      }
      case 'deleteColumns': {
        const entries = args.state.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count)
        const transform = structuralTransformForOp(op)
        return [
          ...clearStructuralSheetMetadataOps(args.state.workbook, op.sheetName, transform),
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
          ...captureStructuralWorkbookMetadataOps(args.state.workbook),
        ]
      }
      case 'setCellValue':
      case 'setCellFormula':
      case 'clearCell':
        return args.restoreCellOps(op.sheetName, op.address)
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

  const buildInverseRecord = (ops: readonly EngineOp[]): TransactionRecord => {
    if (ops.length === 1 && (ops[0]?.kind === 'deleteRows' || ops[0]?.kind === 'deleteColumns')) {
      return buildStructuralDeleteInverseRecord(ops[0])
    }
    return { kind: 'ops', ops: buildInverseOps(ops), potentialNewCells: ops.length }
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

  const canonicalizeStructuralInsertForwardOp = (
    op: Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }>,
  ): Extract<EngineOp, { kind: 'insertRows' | 'insertColumns' }> => {
    if (op.kind === 'insertRows') {
      return {
        ...op,
        entries: args.state.workbook.snapshotRowAxisEntries(op.sheetName, op.start, op.count),
      }
    }
    return {
      ...op,
      entries: args.state.workbook.snapshotColumnAxisEntries(op.sheetName, op.start, op.count),
    }
  }

  const executeTransactionNow = (record: TransactionRecord, source: 'local' | 'restore' | 'undo' | 'redo'): void => {
    if ((record.kind === 'ops' && record.ops.length === 0) || (record.kind === 'cell-mutations' && record.refs.length === 0)) {
      return
    }
    if (record.kind === 'single-existing-numeric-cell-mutation') {
      const ref = singleExistingNumericCellMutationRecordToRef(record)
      const refs = [ref]
      const batch = shouldCreateLocalBatch()
        ? createBatch(args.state.replicaState, [cellMutationRefToEngineOp(args.state.workbook, ref)])
        : null
      args.applyCellMutationsAtBatchNow(refs, batch, source, record.potentialNewCells)
      return
    }
    if (record.kind === 'cell-mutations') {
      const batch = shouldCreateLocalBatch()
        ? createBatch(
            args.state.replicaState,
            record.refs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref)),
          )
        : null
      args.applyCellMutationsAtBatchNow(record.refs, batch, source, record.potentialNewCells)
      return
    }
    const batch = createBatch(args.state.replicaState, record.kind === 'single-op' ? [record.op] : record.ops)
    args.applyBatchNow(batch, source, record.potentialNewCells, record.kind === 'ops' ? record.preparedCellAddressesByOpIndex : undefined)
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
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      isMutationStructuralInsertOp(ops[0]!)
    ) {
      const op = ops[0]
      applyForward(potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells })
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord(canonicalizeStructuralInsertForwardOp(op), potentialNewCells),
          inverse: createLazySingleOpTransactionRecord(inverseMutationStructuralInsertOp(op)),
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      ops[0]?.kind === 'renameSheet'
    ) {
      const op = ops[0]
      applyForward(potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells })
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord(
            { kind: 'renameSheet', oldName: op.oldName, newName: op.newName },
            potentialNewCells,
          ),
          inverse: createLazySingleOpTransactionRecord({ kind: 'renameSheet', oldName: op.newName, newName: op.oldName }),
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    if (
      options.returnUndoOps === false &&
      ops.length === 1 &&
      options.preparedCellAddressesByOpIndex === undefined &&
      ops[0]?.kind === 'upsertDefinedName'
    ) {
      const op = ops[0]
      const existing = args.state.workbook.getDefinedName(op.name)
      const inverseOp: EngineOp =
        existing === undefined
          ? { kind: 'deleteDefinedName', name: op.name }
          : { kind: 'upsertDefinedName', name: existing.name, value: structuredClone(existing.value) }
      applyForward(potentialNewCells === undefined ? { kind: 'single-op', op } : { kind: 'single-op', op, potentialNewCells })
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazySingleOpTransactionRecord({
            kind: 'upsertDefinedName',
            name: op.name,
            value: structuredClone(op.value),
          }),
          inverse: createLazySingleOpTransactionRecord(inverseOp),
        })
        args.state.redoStack.length = 0
      }
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
    const fastHistory =
      tryBuildSingleCellOpHistoryWithoutSnapshot(
        ops,
        potentialNewCells,
        options.returnUndoOps !== false,
        options.reuseForwardOps !== true,
      ) ?? tryBuildFastMutationHistory(fastHistoryArgs)
    const inverse: TransactionRecord = fastHistory?.inverse ?? buildInverseRecord(ops)
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
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(args.state.workbook, inverse)
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
    const shouldCreateBatch = shouldCreateLocalBatch()
    if (options.returnUndoOps === false) {
      const inverse = tryCreateSingleExistingNumericInverseCellMutationRecord(nextRefs) ?? createLazyInverseCellMutationRecord(nextRefs)
      const batch = shouldCreateBatch
        ? createBatch(
            args.state.replicaState,
            nextRefs.map((ref) => cellMutationRefToEngineOp(args.state.workbook, ref)),
          )
        : null
      args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
      if (args.state.getTransactionReplayDepth() === 0) {
        args.state.undoStack.push({
          forward: createLazyCellMutationTransactionRecord(nextRefs, nextPotentialNewCells),
          inverse,
        })
        args.state.redoStack.length = 0
      }
      return null
    }
    const fastHistory = buildFastMutationHistoryFromRefs(nextRefs, nextPotentialNewCells)
    const inverse: TransactionRecord = fastHistory?.inverse ?? {
      kind: 'ops',
      ops: buildInverseOps(transactionRecordOps(args.state.workbook, fastHistory.forward)),
      potentialNewCells: transactionRecordOps(args.state.workbook, fastHistory.forward).length,
    }
    const batch = shouldCreateBatch
      ? createBatch(args.state.replicaState, [...transactionRecordOps(args.state.workbook, fastHistory.forward)])
      : null
    args.applyCellMutationsAtBatchNow(nextRefs, batch, 'local', nextPotentialNewCells)
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: fastHistory?.forward ?? {
          kind: 'ops',
          ops: canonicalizeForwardOps([...transactionRecordOps(args.state.workbook, fastHistory.forward)]),
          potentialNewCells: nextPotentialNewCells,
        },
        inverse,
      })
      args.state.redoStack.length = 0
    }
    return fastHistory?.undoOps ?? cloneTransactionRecordOps(args.state.workbook, inverse)
  }

  const executeLocalExistingNumericCellMutationAtNow = (
    request: EngineExistingNumericCellMutationRef,
    options: {
      returnUndoOps?: boolean
    } = {},
  ): EngineExistingNumericCellMutationResult | null => {
    if (options.returnUndoOps !== false || shouldCreateLocalBatch()) {
      return null
    }
    const cellStore = args.state.workbook.cellStore
    const cellIndex = request.cellIndex
    const oldNumericValue =
      request.trustedExistingNumericLiteral && request.oldNumericValue !== undefined
        ? request.oldNumericValue
        : (cellStore.numbers[cellIndex] ?? 0)
    const result = args.applyExistingNumericCellMutationAtNow?.(request)
    if (!result) {
      return null
    }
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: createSingleExistingNumericCellMutationTransactionRecord(request, 0),
        inverse: createSingleExistingNumericCellMutationTransactionRecord(
          {
            sheetId: request.sheetId,
            cellIndex,
            row: request.row,
            col: request.col,
            value: oldNumericValue,
          },
          1,
        ),
      })
      args.state.redoStack.length = 0
    }
    return result
  }

  const executeLocalExistingLiteralCellMutationAtNow = (
    request: EngineExistingLiteralCellMutationRef,
    options: {
      returnUndoOps?: boolean
    } = {},
  ): EngineExistingNumericCellMutationResult | null => {
    if (typeof request.value === 'number') {
      return executeLocalExistingNumericCellMutationAtNow(
        {
          sheetId: request.sheetId,
          row: request.row,
          col: request.col,
          cellIndex: request.cellIndex,
          value: request.value,
          ...(request.emitTracked === undefined ? {} : { emitTracked: request.emitTracked }),
        },
        options,
      )
    }
    if (options.returnUndoOps !== false || shouldCreateLocalBatch()) {
      return null
    }
    const ref: EngineCellMutationRef = {
      sheetId: request.sheetId,
      cellIndex: request.cellIndex,
      mutation: {
        kind: 'setCellValue',
        row: request.row,
        col: request.col,
        value: request.value,
      },
    }
    const inverse = createLazyInverseCellMutationRecord([ref])
    const result = args.applyExistingLiteralCellMutationAtNow?.(request)
    if (!result) {
      return null
    }
    if (args.state.getTransactionReplayDepth() === 0) {
      args.state.undoStack.push({
        forward: createLazyCellMutationTransactionRecord([ref], 0),
        inverse,
      })
      args.state.redoStack.length = 0
    }
    return result
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

  const executeLocalNowPublic = (
    ops: EngineOp[],
    potentialNewCells?: number,
    options: { readonly returnUndoOps?: boolean } = {},
  ): readonly EngineOp[] | null => {
    if (!shouldCreateLocalBatch()) {
      const refs = tryCellMutationRefsFromOps(ops)
      if (refs !== null) {
        return executeLocalCellMutationsAtNow(refs, potentialNewCells, {
          returnUndoOps: options.returnUndoOps ?? true,
          reuseRefs: true,
        })
      }
    }
    return executeLocalNowWithCustomApply(
      ops,
      potentialNewCells,
      (forward) => {
        executeTransactionNow(forward, 'local')
      },
      { returnUndoOps: options.returnUndoOps ?? true, reuseForwardOps: false },
    )
  }

  const executeLocal = (
    ops: EngineOp[],
    potentialNewCells?: number,
    options: { readonly returnUndoOps?: boolean } = {},
  ): Effect.Effect<readonly EngineOp[] | null, EngineMutationError> =>
    Effect.try({
      try: () => executeLocalNowPublic(ops, potentialNewCells, options),
      catch: (cause) =>
        new EngineMutationError({
          message: 'Failed to execute local transaction',
          cause,
        }),
    })

  const rangeOperations = createMutationRangeOperations({
    workbook: args.state.workbook,
    getCellByIndex: args.getCellByIndex,
    readRangeCells: args.readRangeCells,
    toCellStateOps: args.toCellStateOps,
    executeLocal,
  })

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
    executeLocalNow: executeLocalNowPublic,
    executeLocalCellMutationsAtNow(refs, potentialNewCells) {
      return executeLocalCellMutationsAtNow(refs, potentialNewCells)
    },
    executeLocalExistingNumericCellMutationAtNow(request, options = {}) {
      return executeLocalExistingNumericCellMutationAtNow(request, options)
    },
    executeLocalExistingLiteralCellMutationAtNow(request, options = {}) {
      return executeLocalExistingLiteralCellMutationAtNow(request, options)
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
    executeLocal,
    applyOpsNow(ops, options = {}) {
      const nextOps = options.trusted ? Array.from(ops) : structuredClone([...ops])
      if (nextOps.length === 0) {
        return null
      }
      if (options.captureUndo) {
        return options.returnUndoOps === undefined
          ? this.executeLocalNow(nextOps, options.potentialNewCells)
          : this.executeLocalNow(nextOps, options.potentialNewCells, { returnUndoOps: options.returnUndoOps })
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
              undoOps: cloneTransactionRecordOps(args.state.workbook, inverse),
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
    setRangeValues: rangeOperations.setRangeValues,
    setRangeFormulas: rangeOperations.setRangeFormulas,
    clearRange: rangeOperations.clearRange,
    fillRange: rangeOperations.fillRange,
    copyRange: rangeOperations.copyRange,
    moveRange: rangeOperations.moveRange,
    importSheetCsv: rangeOperations.importSheetCsv,
    renderCommit(ops) {
      return Effect.flatMap(
        Effect.try({
          try: () => {
            if (
              tryExecuteMutationRenderCommitFastPath({
                state: args.state,
                ops,
                hasExternallyVisibleBatchRequirement,
                restoreCellOpFromRef,
                executeLocalNowWithCustomApply,
                executeTransactionNow,
                applyCellMutationsAtNow,
              })
            ) {
              return null
            }
            return normalizeRenderCommitOps(ops)
          },
          catch: (cause) =>
            new EngineMutationError({
              message: 'Failed to normalize render commit operations',
              cause,
            }),
        }),
        (normalized) => {
          if (normalized === null) {
            return Effect.void
          }
          const { engineOps, potentialNewCells, preparedCellAddressesByOpIndex } = normalized
          return Effect.try({
            try: () => {
              executeLocalNowWithCustomApply(
                engineOps,
                potentialNewCells,
                (forward) => {
                  const batchOps = [...transactionRecordOps(args.state.workbook, forward)]
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
          })
        },
      ).pipe(Effect.asVoid)
    },
  }
}
