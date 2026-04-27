import { describe, expect, it } from 'vitest'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
import {
  axisOverridesFromSortedSizes,
  createGridRuntimeAxisOverrideCache,
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
})
