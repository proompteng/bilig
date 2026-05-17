import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot, type RecalcMetrics } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import { ProjectedViewportCellCache } from '../projected-viewport-cell-cache.js'
import { applyProjectedViewportPatch, type ProjectedViewportPatchState } from '../projected-viewport-patch-application.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createPatchState(snapshot: CellSnapshot): ProjectedViewportPatchState {
  const key = `${snapshot.sheetName}!${snapshot.address}`
  return {
    cellSnapshots: new Map([[key, snapshot]]),
    cellKeysBySheet: new Map([[snapshot.sheetName, new Set([key])]]),
    cellStyles: new Map([['style-0', { id: 'style-0' }]]),
    columnSizesBySheet: new Map(),
    columnWidthsBySheet: new Map(),
    pendingColumnWidthsBySheet: new Map(),
    pendingHiddenColumnsBySheet: new Map(),
    rowSizesBySheet: new Map(),
    rowHeightsBySheet: new Map(),
    pendingRowHeightsBySheet: new Map(),
    pendingHiddenRowsBySheet: new Map(),
    hiddenColumnsBySheet: new Map(),
    hiddenRowsBySheet: new Map(),
    freezeRowsBySheet: new Map(),
    freezeColsBySheet: new Map(),
    mergeRangesBySheet: new Map(),
    knownSheets: new Set([snapshot.sheetName]),
  }
}

function fullPatchOmittingB2(): ViewportPatch {
  return {
    version: 12,
    authoritativeRevision: 12,
    full: true,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 1,
      colStart: 1,
      colEnd: 1,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [],
    columns: [],
    rows: [],
  }
}

function clearSnapshot(overrides: Partial<CellSnapshot> = {}): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address: 'B2',
    value: { tag: ValueTag.Empty },
    flags: 0,
    version: 9,
    ...overrides,
  }
}

function textSnapshot(value: string, version = 9): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address: 'B2',
    input: value,
    value: { tag: ValueTag.String, value, stringId: 1 },
    flags: 0,
    version,
  }
}

describe('projected viewport clear tombstones', () => {
  it('keeps confirmed clears through forced same-version selection hydration', () => {
    const cache = new ProjectedViewportCellCache()
    const listener = vi.fn()
    cache.subscribeCells('Sheet1', ['B2'], listener)
    cache.setCellSnapshot(clearSnapshot())
    listener.mockClear()

    expect(cache.setCellSnapshot(textSnapshot('stale-after-clear'), { force: true, forceOptimistic: true })).toBe(false)

    expect(cache.getCell('Sheet1', 'B2')).toEqual(clearSnapshot())
    expect(listener).not.toHaveBeenCalled()
  })

  it('accepts genuinely newer content after a confirmed clear', () => {
    const cache = new ProjectedViewportCellCache()
    cache.setCellSnapshot(clearSnapshot())

    expect(cache.setCellSnapshot(textSnapshot('new-after-clear', 10), { force: true, forceOptimistic: true })).toBe(true)

    expect(cache.getCell('Sheet1', 'B2')).toEqual(textSnapshot('new-after-clear', 10))
  })

  it('keeps an empty tombstone when an authoritative full patch omits a confirmed clear', () => {
    const state = createPatchState(clearSnapshot())

    const result = applyProjectedViewportPatch({
      state,
      patch: fullPatchOmittingB2(),
    })

    expect(result.changedKeys.size).toBe(0)
    expect(result.damage).toEqual([])
    expect(state.cellKeysBySheet.get('Sheet1')?.has('Sheet1!B2')).toBe(true)
    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual(clearSnapshot())
  })

  it('normalizes omitted optimistic clears into stable empty tombstones', () => {
    const state = createPatchState(
      clearSnapshot({
        flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
        styleId: 'stale-style',
        format: 'stale-format',
        numberFormatId: 'stale-number-format',
      }),
    )

    const result = applyProjectedViewportPatch({
      state,
      patch: fullPatchOmittingB2(),
    })

    expect(result.changedKeys.has('Sheet1!B2')).toBe(true)
    expect(result.damage).toEqual([{ cell: [1, 1] }])
    expect(state.cellSnapshots.get('Sheet1!B2')).toEqual(clearSnapshot())
  })
})
