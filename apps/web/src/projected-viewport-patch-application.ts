import { parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellSnapshot, type CellStyleRecord, type Viewport } from '@bilig/protocol'
import type { ViewportPatch } from '@bilig/worker-transport'
import { applyProjectedViewportAxisPatches } from './projected-viewport-axis-patches.js'

export type ProjectedViewportCellItem = readonly [number, number]

export interface ProjectedViewportPatchState {
  readonly cellSnapshots: Map<string, CellSnapshot>
  readonly cellKeysBySheet: Map<string, Set<string>>
  readonly cellStyles: Map<string, CellStyleRecord>
  readonly columnSizesBySheet: Map<string, Record<number, number>>
  readonly columnWidthsBySheet: Map<string, Record<number, number>>
  readonly pendingColumnWidthsBySheet: Map<string, Record<number, number>>
  readonly rowSizesBySheet: Map<string, Record<number, number>>
  readonly rowHeightsBySheet: Map<string, Record<number, number>>
  readonly pendingRowHeightsBySheet: Map<string, Record<number, number>>
  readonly hiddenColumnsBySheet: Map<string, Record<number, true>>
  readonly hiddenRowsBySheet: Map<string, Record<number, true>>
  readonly freezeRowsBySheet: Map<string, number>
  readonly freezeColsBySheet: Map<string, number>
  readonly knownSheets: Set<string>
}

export interface ProjectedViewportPatchApplicationResult {
  readonly damage: readonly { cell: ProjectedViewportCellItem }[]
  readonly changedKeys: ReadonlySet<string>
  readonly axisChanged: boolean
  readonly columnsChanged: boolean
  readonly rowsChanged: boolean
  readonly freezeChanged: boolean
}

function snapshotValueKey(snapshot: CellSnapshot): string {
  switch (snapshot.value.tag) {
    case ValueTag.Number:
      return `n:${snapshot.value.value}`
    case ValueTag.Boolean:
      return `b:${snapshot.value.value ? 1 : 0}`
    case ValueTag.String:
      return `s:${snapshot.value.stringId}:${snapshot.value.value}`
    case ValueTag.Error:
      return `e:${snapshot.value.code}`
    case ValueTag.Empty:
      return 'empty'
  }
  return 'empty'
}

function isResetEmptySnapshot(snapshot: CellSnapshot): boolean {
  return (
    snapshot.value.tag === ValueTag.Empty &&
    snapshot.version === 0 &&
    snapshot.flags === 0 &&
    snapshot.formula === undefined &&
    snapshot.input === undefined &&
    snapshot.format === undefined &&
    snapshot.styleId === undefined &&
    snapshot.numberFormatId === undefined
  )
}

export function cellSnapshotSignature(snapshot: CellSnapshot): string {
  return [
    snapshot.version,
    snapshot.flags,
    snapshot.formula ?? '',
    snapshot.format ?? '',
    snapshot.styleId ?? '',
    snapshot.numberFormatId ?? '',
    snapshot.input ?? '',
    snapshotValueKey(snapshot),
  ].join('|')
}

export function shouldKeepCurrentSnapshot(
  current: CellSnapshot,
  incoming: CellSnapshot,
  options: { readonly allowResetEmptyOverride?: boolean } = {},
): boolean {
  const incomingIsResetEmptySnapshot = isResetEmptySnapshot(incoming)
  const allowResetEmptyOverride = options.allowResetEmptyOverride ?? true
  if (
    incomingIsResetEmptySnapshot &&
    current.version > incoming.version &&
    (current.formula !== undefined || current.input !== undefined)
  ) {
    return true
  }
  if (
    !incomingIsResetEmptySnapshot &&
    incoming.value.tag !== ValueTag.Error &&
    current.formula !== undefined &&
    incoming.formula === undefined &&
    incoming.input === undefined
  ) {
    return true
  }
  if (current.version > incoming.version) {
    return allowResetEmptyOverride ? !incomingIsResetEmptySnapshot : true
  }
  if (current.version < incoming.version) {
    return false
  }
  return (
    !incomingIsResetEmptySnapshot &&
    incoming.value.tag !== ValueTag.Error &&
    current.formula !== undefined &&
    incoming.formula === undefined
  )
}

function cellStyleSignature(style: CellStyleRecord): string {
  const fill = style.fill?.backgroundColor ?? ''
  const font = style.font
  const alignment = style.alignment
  const borders = style.borders
  return [
    fill,
    font?.family ?? '',
    font?.size ?? '',
    font?.bold ? 1 : 0,
    font?.italic ? 1 : 0,
    font?.underline ? 1 : 0,
    font?.color ?? '',
    alignment?.horizontal ?? '',
    alignment?.vertical ?? '',
    alignment?.wrap ? 1 : 0,
    alignment?.indent ?? '',
    borders?.top ? `${borders.top.style}:${borders.top.weight}:${borders.top.color}` : '',
    borders?.right ? `${borders.right.style}:${borders.right.weight}:${borders.right.color}` : '',
    borders?.bottom ? `${borders.bottom.style}:${borders.bottom.weight}:${borders.bottom.color}` : '',
    borders?.left ? `${borders.left.style}:${borders.left.weight}:${borders.left.color}` : '',
  ].join('|')
}

function isCellInsideViewport(snapshot: CellSnapshot, viewport: Viewport): boolean {
  const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
  return (
    parsed.row >= viewport.rowStart && parsed.row <= viewport.rowEnd && parsed.col >= viewport.colStart && parsed.col <= viewport.colEnd
  )
}

function getSheetCellKeys(state: ProjectedViewportPatchState, sheetName: string): Set<string> {
  const existing = state.cellKeysBySheet.get(sheetName)
  if (existing) {
    return existing
  }
  const created = new Set<string>()
  state.cellKeysBySheet.set(sheetName, created)
  return created
}

export function applyProjectedViewportPatch(input: {
  state: ProjectedViewportPatchState
  patch: ViewportPatch
  touchCellKey?: (key: string) => void
}): ProjectedViewportPatchApplicationResult {
  const { state, patch, touchCellKey } = input
  state.knownSheets.add(patch.viewport.sheetName)

  const currentFreezeRows = state.freezeRowsBySheet.get(patch.viewport.sheetName) ?? 0
  const currentFreezeCols = state.freezeColsBySheet.get(patch.viewport.sheetName) ?? 0
  const nextFreezeRows = patch.freezeRows ?? currentFreezeRows
  const nextFreezeCols = patch.freezeCols ?? currentFreezeCols
  const freezeChanged =
    (patch.freezeRows !== undefined || patch.freezeCols !== undefined) &&
    (currentFreezeRows !== nextFreezeRows || currentFreezeCols !== nextFreezeCols)
  if (patch.freezeRows !== undefined || patch.freezeCols !== undefined) {
    state.freezeRowsBySheet.set(patch.viewport.sheetName, nextFreezeRows)
    state.freezeColsBySheet.set(patch.viewport.sheetName, nextFreezeCols)
  }

  const changedKeys = new Set<string>()
  const changedStyleIds = new Set<string>()
  const damagedCellKeys = new Set<string>()
  const damage: { cell: ProjectedViewportCellItem }[] = []

  patch.styles.forEach((style) => {
    const current = state.cellStyles.get(style.id)
    if (!current || cellStyleSignature(current) !== cellStyleSignature(style)) {
      changedStyleIds.add(style.id)
    }
    state.cellStyles.set(style.id, style)
  })

  if (patch.full) {
    const incomingKeys = new Set(patch.cells.map((cell) => `${patch.viewport.sheetName}!${cell.snapshot.address}`))
    const sheetCellKeys = state.cellKeysBySheet.get(patch.viewport.sheetName)
    if (sheetCellKeys) {
      for (const key of sheetCellKeys) {
        if (incomingKeys.has(key)) {
          continue
        }
        const snapshot = state.cellSnapshots.get(key)
        if (!snapshot || !isCellInsideViewport(snapshot, patch.viewport)) {
          continue
        }
        state.cellSnapshots.delete(key)
        sheetCellKeys.delete(key)
        changedKeys.add(key)
        if (!damagedCellKeys.has(key)) {
          const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
          damage.push({ cell: [parsed.col, parsed.row] })
          damagedCellKeys.add(key)
        }
      }
    }
  }

  for (const cell of patch.cells) {
    const key = `${patch.viewport.sheetName}!${cell.snapshot.address}`
    const current = state.cellSnapshots.get(key)
    if (current) {
      const incoming = cell.snapshot
      if (shouldKeepCurrentSnapshot(current, incoming)) {
        continue
      }
      if (cellSnapshotSignature(current) === cellSnapshotSignature(incoming)) {
        if (incoming.styleId && changedStyleIds.has(incoming.styleId) && !damagedCellKeys.has(key)) {
          damage.push({ cell: [cell.col, cell.row] })
          damagedCellKeys.add(key)
        }
        continue
      }
    }

    state.cellSnapshots.set(key, cell.snapshot)
    touchCellKey?.(key)
    getSheetCellKeys(state, patch.viewport.sheetName).add(key)
    changedKeys.add(key)
    if (!damagedCellKeys.has(key)) {
      damage.push({ cell: [cell.col, cell.row] })
      damagedCellKeys.add(key)
    }
  }

  let axisChanged = false
  let columnsChanged = false
  if (patch.columns.length > 0) {
    const nextColumns = applyProjectedViewportAxisPatches({
      patches: patch.columns,
      sizes: state.columnSizesBySheet.get(patch.viewport.sheetName) ?? {},
      renderedSizes: state.columnWidthsBySheet.get(patch.viewport.sheetName) ?? {},
      pendingSizes: state.pendingColumnWidthsBySheet.get(patch.viewport.sheetName) ?? {},
      hiddenAxes: state.hiddenColumnsBySheet.get(patch.viewport.sheetName) ?? {},
    })
    state.columnSizesBySheet.set(patch.viewport.sheetName, nextColumns.sizes)
    state.columnWidthsBySheet.set(patch.viewport.sheetName, nextColumns.renderedSizes)
    state.pendingColumnWidthsBySheet.set(patch.viewport.sheetName, nextColumns.pendingSizes)
    if (Object.keys(nextColumns.hiddenAxes).length === 0) {
      state.hiddenColumnsBySheet.delete(patch.viewport.sheetName)
    } else {
      state.hiddenColumnsBySheet.set(patch.viewport.sheetName, nextColumns.hiddenAxes)
    }
    columnsChanged = nextColumns.axisChanged
    axisChanged = axisChanged || nextColumns.axisChanged
  }

  let rowsChanged = false
  if (patch.rows.length > 0) {
    const nextRows = applyProjectedViewportAxisPatches({
      patches: patch.rows,
      sizes: state.rowSizesBySheet.get(patch.viewport.sheetName) ?? {},
      renderedSizes: state.rowHeightsBySheet.get(patch.viewport.sheetName) ?? {},
      pendingSizes: state.pendingRowHeightsBySheet.get(patch.viewport.sheetName) ?? {},
      hiddenAxes: state.hiddenRowsBySheet.get(patch.viewport.sheetName) ?? {},
    })
    state.rowSizesBySheet.set(patch.viewport.sheetName, nextRows.sizes)
    state.rowHeightsBySheet.set(patch.viewport.sheetName, nextRows.renderedSizes)
    state.pendingRowHeightsBySheet.set(patch.viewport.sheetName, nextRows.pendingSizes)
    if (Object.keys(nextRows.hiddenAxes).length === 0) {
      state.hiddenRowsBySheet.delete(patch.viewport.sheetName)
    } else {
      state.hiddenRowsBySheet.set(patch.viewport.sheetName, nextRows.hiddenAxes)
    }
    rowsChanged = nextRows.axisChanged
    axisChanged = axisChanged || nextRows.axisChanged
  }

  return {
    damage,
    changedKeys,
    axisChanged,
    columnsChanged,
    rowsChanged,
    freezeChanged,
  }
}
