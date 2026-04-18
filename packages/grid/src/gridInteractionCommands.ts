import { formatAddress } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import { cellToEditorSeed } from './gridCells.js'
import type { GridEngineLike } from './grid-engine.js'
import { createSheetSelection } from './gridSelection.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { resolveKeyboardHeaderContextMenuTarget, type WorkbookGridContextMenuTarget } from './workbookGridContextMenuTarget.js'
import type { EditSelectionBehavior, WorkbookGridSurfaceProps } from './workbookGridSurfaceTypes.js'

export function beginWorkbookGridEdit(input: {
  engine: GridEngineLike
  onBeginEdit: WorkbookGridSurfaceProps['onBeginEdit']
  sheetName: string
  address: string
  seed?: string | undefined
  selectionBehavior?: EditSelectionBehavior | undefined
}): void {
  const { address, engine, onBeginEdit, seed, selectionBehavior = 'caret-end', sheetName } = input
  onBeginEdit(seed ?? cellToEditorSeed(engine.getCell(sheetName, address)), selectionBehavior)
}

export function toggleWorkbookGridBooleanCell(input: {
  engine: GridEngineLike
  onToggleBooleanCell: WorkbookGridSurfaceProps['onToggleBooleanCell']
  sheetName: string
  col: number
  row: number
}): boolean {
  const { col, engine, onToggleBooleanCell, row, sheetName } = input
  if (!onToggleBooleanCell) {
    return false
  }
  const address = formatAddress(row, col)
  const snapshot = engine.getCell(sheetName, address)
  if (snapshot.value.tag !== ValueTag.Boolean) {
    return false
  }
  onToggleBooleanCell(sheetName, address, !snapshot.value.value)
  return true
}

export function openWorkbookGridHeaderContextMenuFromKeyboard(input: {
  hostBounds: { left: number; top: number } | null | undefined
  gridSelection: GridSelection
  selectedCell: Item
  getCellScreenBounds: (col: number, row: number) => Rectangle | undefined
  gridMetrics: {
    rowMarkerWidth: number
    headerHeight: number
  }
  openContextMenuForTarget: (target: WorkbookGridContextMenuTarget) => boolean
}): boolean {
  const { getCellScreenBounds, gridMetrics, gridSelection, hostBounds, openContextMenuForTarget, selectedCell } = input
  if (!hostBounds) {
    return false
  }

  const currentCell = gridSelection.current?.cell ?? selectedCell
  const targetCellBounds =
    gridSelection.rows.length > 0 && gridSelection.columns.length === 0
      ? getCellScreenBounds(currentCell[0], gridSelection.rows.first() ?? currentCell[1])
      : gridSelection.columns.length > 0 && gridSelection.rows.length === 0
        ? getCellScreenBounds(gridSelection.columns.first() ?? currentCell[0], currentCell[1])
        : undefined
  const target = resolveKeyboardHeaderContextMenuTarget({
    gridSelection,
    targetCellBounds,
    hostLeft: hostBounds.left,
    hostTop: hostBounds.top,
    rowMarkerWidth: gridMetrics.rowMarkerWidth,
    headerHeight: gridMetrics.headerHeight,
  })
  if (!target) {
    return false
  }
  return openContextMenuForTarget(target)
}

export function selectEntireWorkbookSheet(input: {
  isEditingCell: boolean
  onCommitEdit: WorkbookGridSurfaceProps['onCommitEdit']
  setGridSelection: (selection: GridSelection) => void
  onSelectionChange: (selection: GridSelection) => void
  focusGrid: () => void
}): void {
  const { focusGrid, isEditingCell, onCommitEdit, onSelectionChange, setGridSelection } = input
  if (isEditingCell) {
    onCommitEdit()
  }
  const nextSelection = createSheetSelection()
  setGridSelection(nextSelection)
  onSelectionChange(nextSelection)
  focusGrid()
}
