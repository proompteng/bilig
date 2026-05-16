import type { EdgeArena } from '../../edge-arena.js'
import { makeRangeEntity } from '../../entity-ids.js'
import type { RangeMaterializer, RetargetedRangeDependencies } from '../../range-registry.js'
import type { StructuralTransaction } from '../structural-transaction.js'
import { syncFormulaBindingRangeDependencyEdges, type FormulaBindingReverseEdgeState } from './formula-binding-reverse-edges.js'

export interface FormulaBindingRangeDependencyUpdater {
  readonly refreshRangeDependenciesNow: (rangeIndices: readonly number[]) => void
  readonly retargetRangeDependenciesNow: (transaction: StructuralTransaction, rangeIndices: readonly number[]) => void
}

export interface FormulaBindingRangeDependencyUpdaterArgs {
  readonly state: {
    readonly workbook: {
      readonly cellStore: {
        readonly formulaIds: { readonly [cellIndex: number]: number | undefined }
      }
    }
    readonly ranges: {
      readonly refresh: (
        rangeIndex: number,
        materializer: RangeMaterializer,
      ) => { oldDependencySources: Uint32Array; newDependencySources: Uint32Array }
      readonly applyStructuralTransaction: (
        transaction: StructuralTransaction,
        rangeIndices: readonly number[],
        materializer: RangeMaterializer,
      ) => RetargetedRangeDependencies[]
    }
  }
  readonly edgeArena: EdgeArena
  readonly reverseState: FormulaBindingReverseEdgeState
  readonly ensureCellTrackedByCoords: (sheetId: number, row: number, col: number) => number
  readonly forEachSheetCell: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => void
  readonly scheduleWasmProgramSync: () => void
}

export function createFormulaBindingRangeDependencyUpdater(
  args: FormulaBindingRangeDependencyUpdaterArgs,
): FormulaBindingRangeDependencyUpdater {
  const makeRangeMaterializer = (): RangeMaterializer => ({
    ensureCell: (sheetId: number, row: number, col: number) => args.ensureCellTrackedByCoords(sheetId, row, col),
    forEachSheetCell: (sheetId: number, fn: (cellIndex: number, row: number, col: number) => void) => args.forEachSheetCell(sheetId, fn),
    isFormulaCell: (cellIndex: number) => (args.state.workbook.cellStore.formulaIds[cellIndex] ?? 0) !== 0,
  })

  const syncRangeDependencyEdges = (
    rangeIndex: number,
    deps: { oldDependencySources: Uint32Array; newDependencySources: Uint32Array },
  ): void => {
    syncFormulaBindingRangeDependencyEdges(args.reverseState, args.edgeArena, makeRangeEntity(rangeIndex), deps)
  }

  return {
    refreshRangeDependenciesNow(rangeIndices) {
      const refreshed = new Set<number>()
      const materializer = makeRangeMaterializer()
      rangeIndices.forEach((rangeIndex) => {
        if (refreshed.has(rangeIndex)) {
          return
        }
        refreshed.add(rangeIndex)
        syncRangeDependencyEdges(rangeIndex, args.state.ranges.refresh(rangeIndex, materializer))
      })
      if (refreshed.size > 0) {
        args.scheduleWasmProgramSync()
      }
    },
    retargetRangeDependenciesNow(transaction, rangeIndices) {
      const touched = args.state.ranges.applyStructuralTransaction(transaction, rangeIndices, makeRangeMaterializer())
      touched.forEach(({ rangeIndex, oldDependencySources, newDependencySources }) => {
        syncRangeDependencyEdges(rangeIndex, { oldDependencySources, newDependencySources })
      })
      if (touched.length > 0) {
        args.scheduleWasmProgramSync()
      }
    },
  }
}
