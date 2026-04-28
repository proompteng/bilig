import type { AxisEntryOverride } from '../gridAxisIndex.js'
import { createGridAxisWorldIndexFromRecords, type GridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import type { GridMetrics } from '../gridMetrics.js'
import { resolveGridScrollSpacerSize } from '../gridScrollSurface.js'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

export type SortedAxisSizes = readonly (readonly [number, number])[]

export interface GridRuntimeAxisOverrideCache {
  columnOverrides: readonly AxisEntryOverride[] | null
  rowOverrides: readonly AxisEntryOverride[] | null
  columnSeq: number | null
  rowSeq: number | null
}

export interface GridRuntimeGeometryAxesInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly controlledHiddenColumns?: Readonly<Record<number, true>> | undefined
  readonly controlledHiddenRows?: Readonly<Record<number, true>> | undefined
  readonly freezeCols: number
  readonly freezeRows: number
  readonly gridMetrics: GridMetrics
  readonly hostHeight: number
  readonly hostWidth: number
  readonly rowHeights: Readonly<Record<number, number>>
}

export interface GridRuntimeGeometryAxesState {
  readonly columnAxis: GridAxisWorldIndex
  readonly columnWidthOverridesAttr: string
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly rowAxis: GridAxisWorldIndex
  readonly rowHeightOverridesAttr: string
  readonly runtimeColumnAxisOverrides: readonly AxisEntryOverride[]
  readonly runtimeRowAxisOverrides: readonly AxisEntryOverride[]
  readonly scrollSpacerSize: {
    readonly width: number
    readonly height: number
  }
  readonly sortedColumnWidthOverrides: SortedAxisSizes
  readonly sortedRowHeightOverrides: SortedAxisSizes
}

export function createGridRuntimeAxisOverrideCache(): GridRuntimeAxisOverrideCache {
  return {
    columnOverrides: null,
    rowOverrides: null,
    columnSeq: null,
    rowSeq: null,
  }
}

export function axisOverridesFromSortedSizes(sortedSizes: SortedAxisSizes): readonly AxisEntryOverride[] {
  return sortedSizes.map(([index, size]) => ({ index, size }))
}

export function resolveGridRuntimeGeometryAxes(input: GridRuntimeGeometryAxesInput): GridRuntimeGeometryAxesState {
  const sortedColumnWidthOverrides = sortedAxisSizes(input.columnWidths)
  const sortedRowHeightOverrides = sortedAxisSizes(input.rowHeights)
  const columnAxis = createGridAxisWorldIndexFromRecords({
    axisLength: MAX_COLS,
    defaultSize: input.gridMetrics.columnWidth,
    hidden: input.controlledHiddenColumns,
    sizes: input.columnWidths,
  })
  const rowAxis = createGridAxisWorldIndexFromRecords({
    axisLength: MAX_ROWS,
    defaultSize: input.gridMetrics.rowHeight,
    hidden: input.controlledHiddenRows,
    sizes: input.rowHeights,
  })
  const frozenColumnWidth = columnAxis.span(0, input.freezeCols)
  const frozenRowHeight = rowAxis.span(0, input.freezeRows)

  return {
    columnAxis,
    columnWidthOverridesAttr: stringifyAxisSizes(sortedColumnWidthOverrides),
    frozenColumnWidth,
    frozenRowHeight,
    rowAxis,
    rowHeightOverridesAttr: stringifyAxisSizes(sortedRowHeightOverrides),
    runtimeColumnAxisOverrides: axisOverridesFromSortedSizes(sortedColumnWidthOverrides),
    runtimeRowAxisOverrides: axisOverridesFromSortedSizes(sortedRowHeightOverrides),
    scrollSpacerSize: resolveGridScrollSpacerSize({
      columnAxis,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics: input.gridMetrics,
      hostHeight: input.hostHeight,
      hostWidth: input.hostWidth,
      rowAxis,
    }),
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  }
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

function sortedAxisSizes(sizes: Readonly<Record<number, number>>): SortedAxisSizes {
  return Object.entries(sizes)
    .map(([index, size]) => [Number(index), size] as const)
    .toSorted((left, right) => left[0] - right[0])
}

function stringifyAxisSizes(sortedSizes: SortedAxisSizes): string {
  return sortedSizes.length === 0 ? '{}' : JSON.stringify(Object.fromEntries(sortedSizes))
}
