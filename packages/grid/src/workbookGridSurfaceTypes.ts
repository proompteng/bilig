import type { CellSnapshot, Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import type { GridSelectionSnapshot } from './gridTypes.js'

export type { GridSelectionSnapshot } from './gridTypes.js'

export type EditMovement = readonly [-1 | 0 | 1, -1 | 0 | 1]
export type EditSelectionBehavior = 'select-all' | 'caret-end'
export interface EditTargetSelection {
  readonly sheetName: string
  readonly address: string
}
export type SheetGridViewportSubscription = (
  sheetName: string,
  viewport: Viewport,
  listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
) => () => void

export interface WorkbookGridPreviewRange {
  sheetName: string
  startAddress: string
  endAddress: string
  role: 'target' | 'source'
}

export interface WorkbookGridSurfaceProps {
  engine: GridEngineLike
  sheetName: string
  selectedAddr: string
  selectedCellSnapshot: CellSnapshot
  selectionSnapshot: GridSelectionSnapshot
  editorValue: string
  editorSelectionBehavior: EditSelectionBehavior
  resolvedValue: string
  isEditingCell: boolean
  onSelectionChange(this: void, selection: GridSelectionSnapshot): void
  onSelectionLabelChange?: ((label: string) => void) | undefined
  getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  onBeginEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void
  onEditorChange(this: void, next: string): void
  onCommitEdit(this: void, movement?: EditMovement, valueOverride?: string, targetSelectionOverride?: EditTargetSelection): void
  onCancelEdit(this: void): void
  onClearCell(this: void): void
  onFillRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onCopyRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onMoveRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onToggleBooleanCell?: ((sheetName: string, address: string, nextValue: boolean) => void) | undefined
  onPaste(this: void, sheetName: string, addr: string, values: readonly (readonly string[])[]): void
  subscribeViewport?: SheetGridViewportSubscription | undefined
  columnWidths?: Readonly<Record<number, number>> | undefined
  hiddenColumns?: Readonly<Record<number, true>> | undefined
  hiddenRows?: Readonly<Record<number, true>> | undefined
  rowHeights?: Readonly<Record<number, number>> | undefined
  freezeRows?: number | undefined
  freezeCols?: number | undefined
  onColumnWidthChange?: ((columnIndex: number, newSize: number) => void) | undefined
  onRowHeightChange?: ((rowIndex: number, newSize: number) => void) | undefined
  onSetColumnHidden?: ((columnIndex: number, hidden: boolean) => void) | undefined
  onSetRowHidden?: ((rowIndex: number, hidden: boolean) => void) | undefined
  onInsertRows?: ((startRow: number, count: number) => void) | undefined
  onDeleteRows?: ((startRow: number, count: number) => void) | undefined
  onInsertColumns?: ((startCol: number, count: number) => void) | undefined
  onDeleteColumns?: ((startCol: number, count: number) => void) | undefined
  onSetFreezePane?: ((rows: number, cols: number) => void) | undefined
  onAutofitColumn?: ((columnIndex: number, fallbackWidth: number) => void | Promise<void>) | undefined
  onVisibleViewportChange?: ((viewport: Viewport) => void) | undefined
  previewRanges?: readonly WorkbookGridPreviewRange[] | undefined
  restoreViewportTarget?:
    | {
        readonly token: number
        readonly viewport: Viewport
      }
    | undefined
}
