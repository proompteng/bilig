import { describe, expect, it } from 'vitest'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
import {
  axisOverridesFromSortedSizes,
  createGridRuntimeAxisOverrideCache,
  resolveGridRuntimeGeometryAxes,
  syncGridRuntimeAxisOverrides,
} from '../runtime/gridRuntimeAxisAdapters.js'

const gridMetrics = {
  columnWidth: 100,
  headerHeight: 20,
  rowHeight: 10,
  rowMarkerWidth: 50,
}

describe('grid runtime axis adapters', () => {
  it('converts hook sorted axis sizes and syncs runtime axes only when inputs change', () => {
    const host = new GridRuntimeHost({
      columnCount: 100,
      defaultColumnWidth: 100,
      defaultRowHeight: 10,
      gridMetrics,
      rowCount: 100,
      viewportHeight: 80,
      viewportWidth: 300,
    })
    const cache = createGridRuntimeAxisOverrideCache()
    const columns = axisOverridesFromSortedSizes([
      [0, 160],
      [4, 0],
    ])
    const rows = axisOverridesFromSortedSizes([[2, 30]])

    syncGridRuntimeAxisOverrides(host, cache, {
      columnOverrides: columns,
      columnSeq: 21,
      rowOverrides: rows,
      rowSeq: 22,
    })

    expect(host.snapshot()).toMatchObject({ axisSeqX: 21, axisSeqY: 22 })
    expect(host.columns.sizeAt(0)).toBe(160)
    expect(host.columns.sizeAt(4)).toBe(0)
    expect(host.rows.sizeAt(2)).toBe(30)

    syncGridRuntimeAxisOverrides(host, cache, {
      columnOverrides: columns,
      columnSeq: 21,
      rowOverrides: rows,
      rowSeq: 22,
    })

    expect(host.snapshot()).toMatchObject({ axisSeqX: 21, axisSeqY: 22 })
  })

  it('resolves sorted geometry axes and attrs outside the React hook', () => {
    const state = resolveGridRuntimeGeometryAxes({
      columnWidths: { 4: 140, 1: 80 },
      controlledHiddenColumns: { 3: true },
      controlledHiddenRows: { 2: true },
      freezeCols: 2,
      freezeRows: 1,
      gridMetrics,
      hostHeight: 90,
      hostWidth: 320,
      rowHeights: { 5: 40, 1: 24 },
    })

    expect(state.sortedColumnWidthOverrides).toEqual([
      [1, 80],
      [4, 140],
    ])
    expect(state.sortedRowHeightOverrides).toEqual([
      [1, 24],
      [5, 40],
    ])
    expect(state.columnWidthOverridesAttr).toBe('{"1":80,"4":140}')
    expect(state.rowHeightOverridesAttr).toBe('{"1":24,"5":40}')
    expect(state.runtimeColumnAxisOverrides).toEqual([
      { index: 1, size: 80 },
      { index: 4, size: 140 },
    ])
    expect(state.columnAxis.isHidden(3)).toBe(true)
    expect(state.rowAxis.isHidden(2)).toBe(true)
    expect(state.frozenColumnWidth).toBe(180)
    expect(state.frozenRowHeight).toBe(10)
    expect(state.scrollSpacerSize.width).toBeGreaterThan(0)
    expect(state.scrollSpacerSize.height).toBeGreaterThan(0)
  })
})
