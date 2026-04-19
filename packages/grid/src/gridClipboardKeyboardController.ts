import type { MutableRefObject } from 'react'
import { formatAddress } from '@bilig/formula'
import {
  createColumnSelection,
  createGridSelection,
  createRangeSelection,
  createRowSelection,
  createSheetSelection,
} from './gridSelection.js'
import type { GridSelection, Item } from './gridTypes.js'
import { parseClipboardContent, parseClipboardPlainText } from './gridClipboard.js'
import { cellToEditorSeed } from './gridCells.js'
import { isClipboardShortcut, isHandledGridKey, isNavigationKey, isPrintableKey, normalizeKeyboardKey } from './gridKeyboard.js'
import { resolveGridKeyAction } from './gridKeyActions.js'
import { buildInternalClipboardRange, matchesInternalClipboardPaste, type InternalClipboardRange } from './gridInternalClipboard.js'
import type { GridEngineLike } from './grid-engine.js'
import type { EditMovement, EditSelectionBehavior } from './SheetGridView.js'

export interface GridKeyboardEventLike {
  key: string
  ctrlKey: boolean
  metaKey: boolean
  altKey: boolean
  shiftKey?: boolean
  preventDefault(): void
  cancel?: () => void
}

interface ClipboardDataLike {
  getData(type: string): string
  setData(type: string, value: string): void
}

export interface GridClipboardEventLike {
  clipboardData?: ClipboardDataLike | null
  preventDefault(): void
  stopPropagation(): void
}

interface SelectedCellLike {
  col: number
  row: number
}

interface ApplyGridClipboardValuesOptions {
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  onCopyRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onPaste(this: void, sheetName: string, addr: string, values: readonly (readonly string[])[]): void
  sheetName: string
  target: Item
  values: readonly (readonly string[])[]
}

interface CaptureGridClipboardSelectionOptions {
  engine: GridEngineLike
  gridSelection: GridSelection
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  sheetName: string
}

interface HandleGridKeyOptions {
  applyClipboardValues(this: void, target: Item, values: readonly (readonly string[])[]): void
  beginSelectedEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void
  captureInternalClipboardSelection(this: void): InternalClipboardRange | null
  editorValue: string
  event: GridKeyboardEventLike
  gridSelection: GridSelection
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  isSelectedCellBoolean(this: void): boolean
  isEditingCell: boolean
  onCancelEdit(this: void): void
  onClearCell(this: void): void
  onCommitEdit(this: void, movement?: EditMovement, valueOverride?: string): void
  onEditorChange(this: void, next: string): void
  onSelectionChange(this: void, selection: GridSelection): void
  pendingClipboardCopySequenceRef: MutableRefObject<number>
  pendingKeyboardPasteSequenceRef: MutableRefObject<number>
  pendingTypeSeedRef: MutableRefObject<string | null>
  selectedCell: SelectedCellLike
  setGridSelection(this: void, selection: GridSelection): void
  suppressNextNativePasteRef: MutableRefObject<boolean>
  toggleSelectedBooleanCell(this: void): void
}

interface HandleGridPasteCaptureOptions {
  applyClipboardValues(this: void, target: Item, values: readonly (readonly string[])[]): void
  event: GridClipboardEventLike
  gridSelection: GridSelection
  pendingKeyboardPasteSequenceRef: MutableRefObject<number>
  selectedCell: SelectedCellLike
  suppressNextNativePasteRef: MutableRefObject<boolean>
}

interface HandleGridCopyCaptureOptions {
  captureInternalClipboardSelection(this: void): InternalClipboardRange | null
  event: GridClipboardEventLike
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
}

function isEditableElement(element: EventTarget | null): element is HTMLElement {
  return element instanceof HTMLElement && element.isContentEditable
}

export function isGridKeyboardEditableTarget(target: EventTarget | null): target is HTMLElement {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    isEditableElement(target)
  )
}

function isCellEditorInputFocused(): boolean {
  if (typeof document === 'undefined') {
    return false
  }
  const activeElement = document.activeElement
  return (
    (activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement) &&
    activeElement.dataset['testid'] === 'cell-editor-input'
  )
}

function hasOpenModalDialog(): boolean {
  if (typeof document === 'undefined') {
    return false
  }

  return document.querySelector('[aria-modal="true"]') !== null
}

export function applyGridClipboardValues({
  internalClipboardRef,
  onCopyRange,
  onPaste,
  sheetName,
  target,
  values,
}: ApplyGridClipboardValuesOptions): void {
  if (values.length === 0 || values[0]?.length === 0) {
    return
  }

  const internalClipboard = internalClipboardRef.current
  if (matchesInternalClipboardPaste(internalClipboard, values)) {
    if (!internalClipboard) {
      return
    }
    onCopyRange(
      internalClipboard.sourceStartAddress,
      internalClipboard.sourceEndAddress,
      formatAddress(target[1], target[0]),
      formatAddress(target[1] + internalClipboard.rowCount - 1, target[0] + internalClipboard.colCount - 1),
    )
    return
  }

  onPaste(sheetName, formatAddress(target[1], target[0]), values)
}

export function captureGridClipboardSelection({
  engine,
  gridSelection,
  internalClipboardRef,
  sheetName,
}: CaptureGridClipboardSelectionOptions): InternalClipboardRange | null {
  const range = gridSelection.current?.range
  if (!range || gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
    internalClipboardRef.current = null
    return null
  }

  const values = Array.from({ length: range.height }, (_rowEntry, rowOffset) =>
    Array.from({ length: range.width }, (_colEntry, colOffset) =>
      cellToEditorSeed(engine.getCell(sheetName, formatAddress(range.y + rowOffset, range.x + colOffset))),
    ),
  )

  internalClipboardRef.current = buildInternalClipboardRange(range, values)
  return internalClipboardRef.current
}

export function handleGridKey({
  applyClipboardValues,
  beginSelectedEdit,
  captureInternalClipboardSelection,
  editorValue,
  event,
  gridSelection,
  internalClipboardRef,
  isSelectedCellBoolean,
  isEditingCell,
  onCancelEdit,
  onClearCell,
  onCommitEdit,
  onEditorChange,
  onSelectionChange,
  pendingClipboardCopySequenceRef,
  pendingKeyboardPasteSequenceRef,
  pendingTypeSeedRef,
  selectedCell,
  setGridSelection,
  suppressNextNativePasteRef,
  toggleSelectedBooleanCell,
}: HandleGridKeyOptions): void {
  if (
    !isEditingCell &&
    event.key === ' ' &&
    !event.altKey &&
    !event.ctrlKey &&
    !event.metaKey &&
    !event.shiftKey &&
    isSelectedCellBoolean()
  ) {
    event.preventDefault()
    event.cancel?.()
    toggleSelectedBooleanCell()
    return
  }

  const currentSelectionCell = gridSelection.current?.cell ?? null
  const currentSelectionRange = gridSelection.current?.range ?? null
  const action = resolveGridKeyAction({
    event,
    isEditingCell,
    editorValue,
    editorInputFocused: isCellEditorInputFocused(),
    pendingTypeSeed: pendingTypeSeedRef.current,
    selectedCell: [selectedCell.col, selectedCell.row],
    currentSelectionCell,
    currentRangeAnchor: currentSelectionCell,
    currentSelectionRange,
  })

  if (action.kind === 'none') {
    return
  }

  event.preventDefault()
  event.cancel?.()

  switch (action.kind) {
    case 'edit-append':
      onEditorChange(action.value)
      return
    case 'commit-edit':
      onCommitEdit(action.movement)
      return
    case 'cancel-edit':
      onCancelEdit()
      return
    case 'begin-edit':
      pendingTypeSeedRef.current = action.pendingTypeSeed
      beginSelectedEdit(action.seed, action.selectionBehavior)
      return
    case 'move-selection':
      {
        const nextSelection = createGridSelection(action.cell[0], action.cell[1])
        setGridSelection(nextSelection)
        onSelectionChange(nextSelection)
      }
      return
    case 'extend-selection':
      {
        const nextSelection = createRangeSelection(createGridSelection(action.anchor[0], action.anchor[1]), action.anchor, action.target)
        setGridSelection(nextSelection)
        onSelectionChange(nextSelection)
      }
      return
    case 'clear-cell':
      pendingTypeSeedRef.current = action.pendingTypeSeed
      onClearCell()
      return
    case 'clipboard-copy': {
      const clipboard = captureInternalClipboardSelection()
      if (clipboard && typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
        pendingClipboardCopySequenceRef.current += 1
        const sequence = pendingClipboardCopySequenceRef.current
        void (async () => {
          try {
            await navigator.clipboard.writeText(clipboard.plainText)
          } catch {
            if (pendingClipboardCopySequenceRef.current === sequence) {
              pendingClipboardCopySequenceRef.current = 0
            }
            return
          }
          if (pendingClipboardCopySequenceRef.current === sequence) {
            pendingClipboardCopySequenceRef.current = 0
          }
        })()
      }
      return
    }
    case 'clipboard-cut':
      captureInternalClipboardSelection()
      return
    case 'clipboard-paste': {
      pendingKeyboardPasteSequenceRef.current += 1
      const sequence = pendingKeyboardPasteSequenceRef.current
      if (pendingClipboardCopySequenceRef.current !== 0 && internalClipboardRef.current) {
        if (pendingKeyboardPasteSequenceRef.current !== sequence) {
          return
        }
        pendingKeyboardPasteSequenceRef.current = 0
        applyClipboardValues(action.target, parseClipboardPlainText(internalClipboardRef.current.plainText))
        suppressNextNativePasteRef.current = true
        return
      }
      if (typeof navigator !== 'undefined' && navigator.clipboard?.readText) {
        void (async () => {
          try {
            const rawText = await navigator.clipboard.readText()
            if (pendingKeyboardPasteSequenceRef.current !== sequence) {
              return
            }
            pendingKeyboardPasteSequenceRef.current = 0
            const values = parseClipboardPlainText(rawText)
            applyClipboardValues(action.target, values)
            suppressNextNativePasteRef.current = true
          } catch {
            if (pendingKeyboardPasteSequenceRef.current === sequence) {
              pendingKeyboardPasteSequenceRef.current = 0
            }
          }
        })()
      }
      return
    }
    case 'select-row':
      {
        const nextSelection = createRowSelection(action.col, action.row)
        setGridSelection(nextSelection)
        onSelectionChange(nextSelection)
      }
      return
    case 'select-column':
      {
        const nextSelection = createColumnSelection(action.col, action.row)
        setGridSelection(nextSelection)
        onSelectionChange(nextSelection)
      }
      return
    case 'select-all':
      {
        const nextSelection = createSheetSelection()
        setGridSelection(nextSelection)
        onSelectionChange(nextSelection)
      }
      return
  }
}

export function handleGridCopyCapture({
  captureInternalClipboardSelection,
  event,
  internalClipboardRef,
}: HandleGridCopyCaptureOptions): void {
  captureInternalClipboardSelection()
  if (!event.clipboardData) {
    return
  }
  const clipboard = internalClipboardRef.current
  if (!clipboard) {
    return
  }
  event.clipboardData.setData('text/plain', clipboard.plainText)
  event.preventDefault()
}

export function handleGridPasteCapture({
  applyClipboardValues,
  event,
  gridSelection,
  pendingKeyboardPasteSequenceRef,
  selectedCell,
  suppressNextNativePasteRef,
}: HandleGridPasteCaptureOptions): void {
  if (suppressNextNativePasteRef.current) {
    suppressNextNativePasteRef.current = false
    event.preventDefault()
    event.stopPropagation()
    return
  }
  const rawText = event.clipboardData?.getData('text/plain') ?? ''
  const rawHtml = event.clipboardData?.getData('text/html') ?? ''
  const values = parseClipboardContent(rawText, rawHtml)
  if (values.length === 0 || values[0]?.length === 0) {
    return
  }
  if (pendingKeyboardPasteSequenceRef.current !== 0) {
    pendingKeyboardPasteSequenceRef.current = 0
  }

  const target = gridSelection.current?.cell ?? [selectedCell.col, selectedCell.row]
  applyClipboardValues(target, values)
  event.preventDefault()
  event.stopPropagation()
}

export function shouldHandleGridWindowKey(
  event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>,
  activeElement: Element | null,
  host: HTMLDivElement | null,
): boolean {
  if (hasOpenModalDialog()) {
    return false
  }

  if (isGridKeyboardEditableTarget(activeElement)) {
    return false
  }

  const withinGridHost = Boolean(activeElement && host?.contains(activeElement))
  const onDocumentBody = activeElement === document.body || activeElement === document.documentElement || activeElement === null
  if (withinGridHost) {
    return isHandledGridKey(event)
  }
  if (!onDocumentBody) {
    return false
  }

  return isHandledGridKey(event)
}

export function shouldHandleGridSurfaceKey(event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey'>): boolean {
  return (
    isPrintableKey(event) ||
    isClipboardShortcut(event) ||
    isNavigationKey(event.key) ||
    event.key === 'Enter' ||
    event.key === 'Tab' ||
    event.key === 'F2' ||
    event.key === 'Backspace' ||
    event.key === 'Delete'
  )
}

export function getNormalizedGridKeyboardKey(key: string, code?: string): string {
  return normalizeKeyboardKey(key, code)
}
