import type { AxisEntryOverride } from '../gridAxisIndex.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

export interface GridRuntimeAxisOverrideCache {
  columnOverrides: readonly AxisEntryOverride[] | null
  rowOverrides: readonly AxisEntryOverride[] | null
  columnSeq: number | null
  rowSeq: number | null
}

export function createGridRuntimeAxisOverrideCache(): GridRuntimeAxisOverrideCache {
  return {
    columnOverrides: null,
    rowOverrides: null,
    columnSeq: null,
    rowSeq: null,
  }
}

export function axisOverridesFromSortedSizes(sortedSizes: readonly (readonly [number, number])[]): readonly AxisEntryOverride[] {
  return sortedSizes.map(([index, size]) => ({ index, size }))
}

export function syncGridRuntimeAxisOverrides(
  host: GridRuntimeHost,
  cache: GridRuntimeAxisOverrideCache,
  input: {
    readonly columnOverrides: readonly AxisEntryOverride[]
    readonly rowOverrides: readonly AxisEntryOverride[]
    readonly columnSeq: number
    readonly rowSeq: number
  },
): void {
  const updateColumns = cache.columnOverrides !== input.columnOverrides || cache.columnSeq !== input.columnSeq
  const updateRows = cache.rowOverrides !== input.rowOverrides || cache.rowSeq !== input.rowSeq
  if (!updateColumns && !updateRows) {
    return
  }
  host.updateAxes({
    ...(updateColumns ? { columns: input.columnOverrides, columnSeq: input.columnSeq } : {}),
    ...(updateRows ? { rows: input.rowOverrides, rowSeq: input.rowSeq } : {}),
  })
  cache.columnOverrides = input.columnOverrides
  cache.rowOverrides = input.rowOverrides
  cache.columnSeq = input.columnSeq
  cache.rowSeq = input.rowSeq
}
