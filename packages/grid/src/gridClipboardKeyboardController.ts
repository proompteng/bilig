import type { MutableRefObject } from 'react'
import { formatAddress } from '@bilig/formula'
import {
  createColumnSelection,
  createGridSelection,
  createRectangleSelectionFromRange,
  createRangeSelection,
  createRowSelection,
  createSheetSelection,
  selectionToSnapshot,
} from './gridSelection.js'
import { CompactSelection, type GridSelection, type GridSelectionSnapshot, type Item } from './gridTypes.js'
import { parseClipboardContent, parseClipboardPlainText } from './gridClipboard.js'
import { cellToEditorSeed, snapshotToRenderCell } from './gridCells.js'
import {
  isClipboardShortcut,
  isClearCellKey,
  isCurrentRegionSelectionShortcut,
  isFillSelectionShortcut,
  isHandledGridKey,
  isFillShortcut,
  isNavigationShortcut,
  isPrintableKey,
  isScrollActiveCellShortcut,
  isSheetSelectionShortcut,
  isStructuralDeleteShortcut,
  normalizeKeyboardKey,
} from './gridKeyboard.js'
import { resolveGridKeyAction, type GridAxisDeleteRange } from './gridKeyActions.js'
import { buildInternalClipboardRange, matchesInternalClipboardPaste, type InternalClipboardRange } from './gridInternalClipboard.js'
import type { GridEngineLike } from './grid-engine.js'
import type { GridKeyNavigationResolver } from './gridNavigation.js'
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
  onMoveRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onPaste(this: void, sheetName: string, addr: string, values: readonly (readonly string[])[]): void
  pasteValuesOnly?: boolean | undefined
  sheetName: string
  target: Item
  values: readonly (readonly string[])[]
}

interface CaptureGridClipboardSelectionOptions {
  engine: GridEngineLike
  getCellEditorSeed?: ((sheetName: string, address: string) => string | undefined) | undefined
  gridSelection: GridSelection
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  operation?: InternalClipboardRange['operation']
  sheetName: string
}

interface ApplyClipboardValuesOptions {
  readonly pasteValuesOnly?: boolean | undefined
}

interface HandleGridKeyOptions {
  applyClipboardValues(this: void, target: Item, values: readonly (readonly string[])[], options?: ApplyClipboardValuesOptions): void
  beginSelectedEdit(this: void, seed?: string, selectionBehavior?: EditSelectionBehavior): void
  captureInternalClipboardSelection(this: void, operation?: InternalClipboardRange['operation']): InternalClipboardRange | null
  editorValue: string
  event: GridKeyboardEventLike
  gridSelection: GridSelection
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  isSelectedCellBoolean(this: void): boolean
  isEditingCell: boolean
  onCancelEdit(this: void): void
  onClearCell(this: void, selection?: GridSelectionSnapshot): void
  onCommitEdit(this: void, movement?: EditMovement, valueOverride?: string): void
  onDeleteColumns?: ((startCol: number, count: number) => void | Promise<void>) | undefined
  onDeleteRows?: ((startRow: number, count: number) => void | Promise<void>) | undefined
  onEditorChange(this: void, next: string): void
  onFillRange(this: void, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string): void
  onSelectionChange(this: void, selection: GridSelection): void
  navigation?: GridKeyNavigationResolver | null
  pageJumpRows?: number | null
  scrollActiveCellIntoView(this: void): void
  pendingClipboardCopySequenceRef: MutableRefObject<number>
  pendingKeyboardPasteSequenceRef: MutableRefObject<number>
  pendingTypeSeedRef: MutableRefObject<string | null>
  selectedCell: SelectedCellLike
  setGridSelection(this: void, selection: GridSelection): void
  sheetName: string
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
  captureInternalClipboardSelection(this: void, operation?: InternalClipboardRange['operation']): InternalClipboardRange | null
  event: GridClipboardEventLike
  internalClipboardRef: MutableRefObject<InternalClipboardRange | null>
  operation?: InternalClipboardRange['operation']
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

function selectionRangesToDeleteRanges(ranges: GridSelection['rows']['ranges']): readonly GridAxisDeleteRange[] {
  return ranges.map(([start, endExclusive]) => ({ start, count: endExclusive - start })).filter((range) => range.count > 0)
}

function runStructuralDeleteRanges(
  ranges: readonly GridAxisDeleteRange[],
  onDeleteRange: ((start: number, count: number) => void | Promise<void>) | undefined,
): void {
  if (!onDeleteRange) {
    return
  }

  void Promise.all(ranges.toSorted((left, right) => right.start - left.start).map((range) => onDeleteRange(range.start, range.count)))
}

export function applyGridClipboardValues({
  internalClipboardRef,
  onCopyRange,
  onMoveRange,
  onPaste,
  pasteValuesOnly = false,
  sheetName,
  target,
  values,
}: ApplyGridClipboardValuesOptions): void {
  if (values.length === 0 || values[0]?.length === 0) {
    return
  }

  const internalClipboard = internalClipboardRef.current
  if (!pasteValuesOnly && matchesInternalClipboardPaste(internalClipboard, values)) {
    if (!internalClipboard) {
      return
    }
    const targetStartAddress = formatAddress(target[1], target[0])
    const targetEndAddress = formatAddress(target[1] + internalClipboard.rowCount - 1, target[0] + internalClipboard.colCount - 1)
    if (internalClipboard.operation === 'cut') {
      onMoveRange(internalClipboard.sourceStartAddress, internalClipboard.sourceEndAddress, targetStartAddress, targetEndAddress)
      internalClipboardRef.current = null
    } else {
      onCopyRange(internalClipboard.sourceStartAddress, internalClipboard.sourceEndAddress, targetStartAddress, targetEndAddress)
    }
    return
  }

  onPaste(sheetName, formatAddress(target[1], target[0]), values)
}

function resolveKeyboardClipboardValues(clipboard: InternalClipboardRange, pasteValuesOnly: boolean): readonly (readonly string[])[] {
  return parseClipboardPlainText(pasteValuesOnly ? clipboard.valuesOnlyPlainText : clipboard.plainText)
}

function resolveSystemClipboardValues(
  rawText: string,
  internalClipboard: InternalClipboardRange | null,
  pasteValuesOnly: boolean,
): readonly (readonly string[])[] {
  if (pasteValuesOnly && internalClipboard && rawText === internalClipboard.plainText) {
    return parseClipboardPlainText(internalClipboard.valuesOnlyPlainText)
  }
  return parseClipboardPlainText(rawText)
}

export function captureGridClipboardSelection({
  engine,
  getCellEditorSeed,
  gridSelection,
  internalClipboardRef,
  operation = 'copy',
  sheetName,
}: CaptureGridClipboardSelectionOptions): InternalClipboardRange | null {
  const range = gridSelection.current?.range
  if (!range || gridSelection.columns.length > 0 || gridSelection.rows.length > 0) {
    internalClipboardRef.current = null
    return null
  }

  const values: string[][] = []
  const valuesOnly: string[][] = []

  for (let rowOffset = 0; rowOffset < range.height; rowOffset += 1) {
    const rowValues: string[] = []
    const rowValuesOnly: string[] = []
    for (let colOffset = 0; colOffset < range.width; colOffset += 1) {
      const address = formatAddress(range.y + rowOffset, range.x + colOffset)
      const editorSeed = getCellEditorSeed?.(sheetName, address)
      if (editorSeed !== undefined) {
        rowValues.push(editorSeed)
        rowValuesOnly.push(editorSeed)
        continue
      }
      const snapshot = engine.getCell(sheetName, address)
      rowValues.push(cellToEditorSeed(snapshot))
      rowValuesOnly.push(snapshotToRenderCell(snapshot, engine.getCellStyle(snapshot.styleId)).displayText)
    }
    values.push(rowValues)
    valuesOnly.push(rowValuesOnly)
  }

  internalClipboardRef.current = buildInternalClipboardRange(range, values, operation, valuesOnly)
  return internalClipboardRef.current
}

function writeClipboardPlainTextFromKeyboard(
  clipboard: InternalClipboardRange | null,
  pendingClipboardCopySequenceRef: MutableRefObject<number>,
): void {
  if (!clipboard || typeof navigator === 'undefined' || !navigator.clipboard?.writeText) {
    return
  }
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
  onDeleteColumns,
  onDeleteRows,
  onEditorChange,
  onFillRange,
  onSelectionChange,
  navigation,
  pageJumpRows,
  scrollActiveCellIntoView,
  pendingClipboardCopySequenceRef,
  pendingKeyboardPasteSequenceRef,
  pendingTypeSeedRef,
  selectedCell,
  setGridSelection,
  sheetName,
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
    selectedColumnRanges: selectionRangesToDeleteRanges(gridSelection.columns.ranges),
    selectedRowRanges: selectionRangesToDeleteRanges(gridSelection.rows.ranges),
    navigation,
    pageJumpRows,
  })

  if (action.kind === 'none') {
    if (!isEditingCell && isClearCellKey(event)) {
      event.preventDefault()
      event.cancel?.()
    }
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
    case 'move-selection-in-range':
      {
        const nextSelection: GridSelection = {
          current: {
            cell: action.cell,
            range: { ...action.range },
            rangeStack: gridSelection.current?.rangeStack ?? [],
          },
          columns: CompactSelection.empty(),
          rows: CompactSelection.empty(),
        }
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
      onClearCell(selectionToSnapshot(gridSelection, sheetName, formatAddress(selectedCell.row, selectedCell.col)))
      return
    case 'scroll-active-cell':
      scrollActiveCellIntoView()
      return
    case 'clipboard-copy': {
      writeClipboardPlainTextFromKeyboard(captureInternalClipboardSelection('copy'), pendingClipboardCopySequenceRef)
      return
    }
    case 'clipboard-cut': {
      writeClipboardPlainTextFromKeyboard(captureInternalClipboardSelection('cut'), pendingClipboardCopySequenceRef)
      return
    }
    case 'clipboard-paste': {
      pendingKeyboardPasteSequenceRef.current += 1
      const sequence = pendingKeyboardPasteSequenceRef.current
      if (pendingClipboardCopySequenceRef.current !== 0 && internalClipboardRef.current) {
        if (pendingKeyboardPasteSequenceRef.current !== sequence) {
          return
        }
        pendingKeyboardPasteSequenceRef.current = 0
        applyClipboardValues(action.target, resolveKeyboardClipboardValues(internalClipboardRef.current, action.valuesOnly), {
          pasteValuesOnly: action.valuesOnly,
        })
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
            const values = resolveSystemClipboardValues(rawText, internalClipboardRef.current, action.valuesOnly)
            applyClipboardValues(action.target, values, { pasteValuesOnly: action.valuesOnly })
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
    case 'fill-range':
      onFillRange(
        formatAddress(action.source.y, action.source.x),
        formatAddress(action.source.y + action.source.height - 1, action.source.x + action.source.width - 1),
        formatAddress(action.target.y, action.target.x),
        formatAddress(action.target.y + action.target.height - 1, action.target.x + action.target.width - 1),
      )
      return
    case 'delete-selected-rows':
      runStructuralDeleteRanges(action.ranges, onDeleteRows)
      return
    case 'delete-selected-columns':
      runStructuralDeleteRanges(action.ranges, onDeleteColumns)
      return
    case 'handled':
      return
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
    case 'select-range':
      {
        const nextSelection = {
          ...createRectangleSelectionFromRange(action.range),
          current: {
            cell: action.cell,
            range: { ...action.range },
            rangeStack: [],
          },
        }
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
  operation = 'copy',
}: HandleGridCopyCaptureOptions): void {
  captureInternalClipboardSelection(operation)
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
  const workbookScope = host?.closest('[data-workbook-keyboard-scope="true"]') ?? null
  const withinWorkbookChrome = Boolean(activeElement && workbookScope?.contains(activeElement))
  if (withinGridHost) {
    return isHandledGridKey(event)
  }
  if (withinWorkbookChrome) {
    return isWorkbookChromeGridShortcut(event)
  }
  if (!onDocumentBody) {
    return false
  }

  return isHandledGridKey(event)
}

export function shouldSuppressWorkbookChromeClearKey(
  event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>,
  activeElement: Element | null,
  host: HTMLDivElement | null,
): boolean {
  if (hasOpenModalDialog() || isGridKeyboardEditableTarget(activeElement) || !isClearCellKey(event)) {
    return false
  }

  const workbookScope = host?.closest('[data-workbook-keyboard-scope="true"]') ?? null
  return Boolean(activeElement && !host?.contains(activeElement) && workbookScope?.contains(activeElement))
}

export function shouldSuppressWorkbookChromeSelectionKeyUp(
  event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>,
  activeElement: Element | null,
  host: HTMLDivElement | null,
): boolean {
  if (hasOpenModalDialog() || isGridKeyboardEditableTarget(activeElement) || !isSheetSelectionShortcut(event)) {
    return false
  }

  const workbookScope = host?.closest('[data-workbook-keyboard-scope="true"]') ?? null
  return Boolean(activeElement && !host?.contains(activeElement) && workbookScope?.contains(activeElement))
}

function isWorkbookChromeGridShortcut(event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>): boolean {
  const hasPrimaryModifier = event.ctrlKey || event.metaKey
  const normalizedKey = event.key.toLowerCase()
  return (
    isClipboardShortcut(event) ||
    isFillShortcut(event) ||
    isFillSelectionShortcut(event) ||
    isScrollActiveCellShortcut(event) ||
    isStructuralDeleteShortcut(event) ||
    isNavigationShortcut(event) ||
    isCurrentRegionSelectionShortcut(event) ||
    (hasPrimaryModifier && !event.altKey && normalizedKey === 'a') ||
    isSheetSelectionShortcut(event) ||
    (!event.altKey && !hasPrimaryModifier && !event.shiftKey && event.key === 'F2') ||
    (!event.altKey && (event.key === 'Home' || event.key === 'End' || event.key === 'PageUp' || event.key === 'PageDown'))
  )
}

export function shouldHandleGridSurfaceKey(
  event: Pick<GridKeyboardEventLike, 'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'>,
): boolean {
  return (
    isPrintableKey(event) ||
    isClipboardShortcut(event) ||
    isFillShortcut(event) ||
    isFillSelectionShortcut(event) ||
    isScrollActiveCellShortcut(event) ||
    isStructuralDeleteShortcut(event) ||
    isNavigationShortcut(event) ||
    isCurrentRegionSelectionShortcut(event) ||
    isSheetSelectionShortcut(event) ||
    ((event.ctrlKey || event.metaKey) && !event.altKey && event.key.toLowerCase() === 'a') ||
    (event.key === 'Enter' && !event.altKey && !event.ctrlKey && !event.metaKey) ||
    (event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey) ||
    (event.key === 'F2' && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) ||
    isClearCellKey(event)
  )
}

export function getNormalizedGridKeyboardKey(key: string, code?: string): string {
  return normalizeKeyboardKey(key, code)
}
