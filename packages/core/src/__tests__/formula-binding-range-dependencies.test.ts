import { describe, expect, it, vi } from 'vitest'
import { EdgeArena } from '../edge-arena.js'
import { makeCellEntity, makeRangeEntity } from '../entity-ids.js'
import {
  appendFormulaBindingReverseEdge,
  getFormulaBindingReverseEdgeSlice,
  type FormulaBindingReverseEdgeState,
} from '../engine/services/formula-binding-reverse-edges.js'
import { createFormulaBindingRangeDependencyUpdater } from '../engine/services/formula-binding-range-dependencies.js'
import { buildStructuralTransaction } from '../engine/structural-transaction.js'

function createReverseState(): FormulaBindingReverseEdgeState {
  return {
    reverseCellEdges: [],
    reverseRangeEdges: [],
    reverseDefinedNameEdges: new Map(),
    reverseTableEdges: new Map(),
    reverseSpillEdges: new Map(),
    reverseAggregateColumnEdges: new Map(),
    reverseExactLookupColumnEdges: new Map(),
    reverseSortedLookupColumnEdges: new Map(),
  }
}

describe('formula binding range dependency updater', () => {
  it('deduplicates refreshed ranges while syncing reverse dependency edges', () => {
    const edgeArena = new EdgeArena()
    const reverseState = createReverseState()
    const rangeEntity = makeRangeEntity(7)
    const staleDependency = makeCellEntity(1)
    const retainedDependency = makeCellEntity(2)
    const newDependency = makeCellEntity(3)
    appendFormulaBindingReverseEdge(reverseState, edgeArena, staleDependency, rangeEntity)
    appendFormulaBindingReverseEdge(reverseState, edgeArena, retainedDependency, rangeEntity)

    const ensureCellTrackedByCoords = vi.fn(() => 10)
    const forEachSheetCell = vi.fn()
    const scheduleWasmProgramSync = vi.fn()
    const refresh = vi.fn((rangeIndex: number, materializer: { isFormulaCell?: (cellIndex: number) => boolean }) => {
      expect(rangeIndex).toBe(7)
      expect(materializer.isFormulaCell?.(4)).toBe(true)
      expect(materializer.isFormulaCell?.(5)).toBe(false)
      return {
        oldDependencySources: Uint32Array.from([staleDependency, retainedDependency]),
        newDependencySources: Uint32Array.from([retainedDependency, newDependency]),
      }
    })
    const updater = createFormulaBindingRangeDependencyUpdater({
      state: {
        workbook: {
          cellStore: { formulaIds: { 4: 22 } },
        },
        ranges: {
          refresh,
          applyStructuralTransaction: vi.fn(() => []),
        },
      },
      edgeArena,
      reverseState,
      ensureCellTrackedByCoords,
      forEachSheetCell,
      scheduleWasmProgramSync,
    })

    updater.refreshRangeDependenciesNow([7, 7])

    expect(refresh).toHaveBeenCalledTimes(1)
    expect(scheduleWasmProgramSync).toHaveBeenCalledTimes(1)
    expect(getFormulaBindingReverseEdgeSlice(reverseState, staleDependency)).toBeUndefined()
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(reverseState, retainedDependency)!)).toEqual(Uint32Array.from([rangeEntity]))
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(reverseState, newDependency)!)).toEqual(Uint32Array.from([rangeEntity]))
  })

  it('only schedules wasm sync for structural retargets that touch ranges', () => {
    const edgeArena = new EdgeArena()
    const reverseState = createReverseState()
    const rangeEntity = makeRangeEntity(9)
    const oldDependency = makeCellEntity(11)
    const nextDependency = makeCellEntity(12)
    appendFormulaBindingReverseEdge(reverseState, edgeArena, oldDependency, rangeEntity)
    const scheduleWasmProgramSync = vi.fn()
    const applyStructuralTransaction = vi
      .fn()
      .mockReturnValueOnce([])
      .mockReturnValueOnce([
        {
          rangeIndex: 9,
          oldDependencySources: Uint32Array.from([oldDependency]),
          newDependencySources: Uint32Array.from([nextDependency]),
        },
      ])
    const updater = createFormulaBindingRangeDependencyUpdater({
      state: {
        workbook: {
          cellStore: { formulaIds: [] },
        },
        ranges: {
          refresh: vi.fn(),
          applyStructuralTransaction,
        },
      },
      edgeArena,
      reverseState,
      ensureCellTrackedByCoords: vi.fn(() => 1),
      forEachSheetCell: vi.fn(),
      scheduleWasmProgramSync,
    })
    const transaction = buildStructuralTransaction({
      sheetName: 'Sheet1',
      sheetId: 1,
      transform: { kind: 'insert', axis: 'row', start: 0, count: 1 },
      remappedCells: [],
    })

    updater.retargetRangeDependenciesNow(transaction, [9])
    updater.retargetRangeDependenciesNow(transaction, [9])

    expect(scheduleWasmProgramSync).toHaveBeenCalledTimes(1)
    expect(getFormulaBindingReverseEdgeSlice(reverseState, oldDependency)).toBeUndefined()
    expect(edgeArena.read(getFormulaBindingReverseEdgeSlice(reverseState, nextDependency)!)).toEqual(Uint32Array.from([rangeEntity]))
  })
})
