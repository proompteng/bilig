import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { clampCell } from './gridSelection.js'
import {
  isClearCellKey,
  isCurrentRegionSelectionShortcut,
  isFillSelectionShortcut,
  isFillShortcut,
  isScrollActiveCellShortcut,
  isStructuralDeleteShortcut,
} from './gridKeyboard.js'
import type { GridKeyNavigationResolver, GridNavigationDirection } from './gridNavigation.js'
import type { Item, Rectangle } from './gridTypes.js'

export type GridEditSelectionBehavior = 'select-all' | 'caret-end'

export interface GridKeyActionEvent {
  key: string
  ctrlKey: boolean
  metaKey: boolean
  altKey: boolean
  shiftKey?: boolean
}

export type GridKeyAction =
  | { kind: 'none' }
  | { kind: 'edit-append'; value: string }
  | { kind: 'commit-edit'; movement?: readonly [-1 | 0 | 1, -1 | 0 | 1] }
  | { kind: 'cancel-edit' }
  | {
      kind: 'begin-edit'
      seed?: string
      selectionBehavior: GridEditSelectionBehavior
      pendingTypeSeed: string | null
    }
  | { kind: 'move-selection'; cell: Item }
  | { kind: 'move-selection-in-range'; cell: Item; range: Rectangle }
  | { kind: 'extend-selection'; anchor: Item; target: Item }
  | { kind: 'clear-cell'; pendingTypeSeed: null }
  | { kind: 'clipboard-copy' }
  | { kind: 'clipboard-cut' }
  | { kind: 'clipboard-paste'; target: Item; valuesOnly: boolean }
  | { kind: 'fill-range'; source: Rectangle; target: Rectangle }
  | { kind: 'delete-selected-rows'; ranges: readonly GridAxisDeleteRange[] }
  | { kind: 'delete-selected-columns'; ranges: readonly GridAxisDeleteRange[] }
  | { kind: 'scroll-active-cell' }
  | { kind: 'handled' }
  | { kind: 'select-row'; row: number; col: number }
  | { kind: 'select-column'; row: number; col: number }
  | { kind: 'select-range'; cell: Item; range: Rectangle }
  | { kind: 'select-all' }

const PAGE_JUMP_ROWS = 20

export interface GridAxisDeleteRange {
  readonly start: number
  readonly count: number
}

function moveSelectionToEdge(cell: Item, direction: 'up' | 'down' | 'left' | 'right'): Item {
  switch (direction) {
    case 'up':
      return [cell[0], 0]
    case 'down':
      return [cell[0], MAX_ROWS - 1]
    case 'left':
      return [0, cell[1]]
    case 'right':
      return [MAX_COLS - 1, cell[1]]
  }
}

function resolveSelectionActiveCell(
  anchorCell: Item,
  currentSelectionCell: Item | null,
  currentSelectionRange: Rectangle | null | undefined,
): Item {
  if (!currentSelectionRange || (currentSelectionRange.width === 1 && currentSelectionRange.height === 1)) {
    return currentSelectionCell ?? anchorCell
  }

  const horizontalTarget =
    anchorCell[0] === currentSelectionRange.x ? currentSelectionRange.x + currentSelectionRange.width - 1 : currentSelectionRange.x
  const verticalTarget =
    anchorCell[1] === currentSelectionRange.y ? currentSelectionRange.y + currentSelectionRange.height - 1 : currentSelectionRange.y

  return [horizontalTarget, verticalTarget]
}

function isMultiCellRange(range: Rectangle | null | undefined): range is Rectangle {
  return Boolean(range && (range.width > 1 || range.height > 1))
}

function isCellInRange(cell: Item, range: Rectangle): boolean {
  return cell[0] >= range.x && cell[0] < range.x + range.width && cell[1] >= range.y && cell[1] < range.y + range.height
}

function resolveRangeNavigationCell(selectedCell: Item, currentSelectionCell: Item | null, range: Rectangle): Item {
  if (currentSelectionCell && isCellInRange(currentSelectionCell, range)) {
    return currentSelectionCell
  }
  if (isCellInRange(selectedCell, range)) {
    return selectedCell
  }
  return [range.x, range.y]
}

function moveWithinRange(cell: Item, range: Rectangle, axis: 'horizontal' | 'vertical', reverse: boolean): Item {
  const total = range.width * range.height
  if (total <= 1) {
    return [range.x, range.y]
  }

  const colOffset = Math.min(range.width - 1, Math.max(0, cell[0] - range.x))
  const rowOffset = Math.min(range.height - 1, Math.max(0, cell[1] - range.y))
  const currentIndex = axis === 'horizontal' ? rowOffset * range.width + colOffset : colOffset * range.height + rowOffset
  const nextIndex = (currentIndex + (reverse ? -1 : 1) + total) % total

  if (axis === 'horizontal') {
    return [range.x + (nextIndex % range.width), range.y + Math.floor(nextIndex / range.width)]
  }

  return [range.x + Math.floor(nextIndex / range.height), range.y + (nextIndex % range.height)]
}

interface ResolveGridKeyActionOptions {
  event: GridKeyActionEvent
  isEditingCell: boolean
  editorValue: string
  editorInputFocused: boolean
  pendingTypeSeed: string | null
  selectedCell: Item
  currentSelectionCell: Item | null
  currentRangeAnchor: Item | null
  currentSelectionRange?: Rectangle | null
  selectedColumnRanges?: readonly GridAxisDeleteRange[] | null | undefined
  selectedRowRanges?: readonly GridAxisDeleteRange[] | null | undefined
  navigation?: GridKeyNavigationResolver | null | undefined
  pageJumpRows?: number | null | undefined
}

function rectanglesEqual(left: Rectangle | null | undefined, right: Rectangle | null | undefined): boolean {
  return Boolean(left && right && left.x === right.x && left.y === right.y && left.width === right.width && left.height === right.height)
}

function resolvePageJumpRows(rawPageJumpRows: number | null | undefined): number {
  if (typeof rawPageJumpRows !== 'number' || !Number.isFinite(rawPageJumpRows)) {
    return PAGE_JUMP_ROWS
  }
  return Math.max(1, Math.floor(rawPageJumpRows))
}

export function resolveGridKeyAction(options: ResolveGridKeyActionOptions): GridKeyAction {
  const {
    event,
    isEditingCell,
    editorValue,
    editorInputFocused,
    pendingTypeSeed,
    selectedCell,
    currentSelectionCell,
    currentRangeAnchor,
    currentSelectionRange,
    selectedColumnRanges,
    selectedRowRanges,
    navigation,
    pageJumpRows,
  } = options

  const anchorCell = currentRangeAnchor ?? selectedCell
  const selectedActionCell = currentSelectionCell ?? selectedCell
  const activeCell = resolveSelectionActiveCell(anchorCell, currentSelectionCell, currentSelectionRange)

  if (isEditingCell) {
    if (!editorInputFocused) {
      if (event.key === 'Enter' && !event.altKey && !event.ctrlKey && !event.metaKey) {
        return { kind: 'commit-edit', movement: [0, event.shiftKey ? -1 : 1] }
      }
      if (event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey) {
        return { kind: 'commit-edit', movement: [event.shiftKey ? -1 : 1, 0] }
      }
      if (event.key === 'Escape' && !event.altKey && !event.ctrlKey && !event.metaKey) {
        return { kind: 'cancel-edit' }
      }
      if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
        return { kind: 'edit-append', value: `${editorValue}${event.key}` }
      }
    }
    return { kind: 'none' }
  }

  if (event.key === 'F2' && !event.altKey && !event.ctrlKey && !event.metaKey && !event.shiftKey) {
    return { kind: 'begin-edit', selectionBehavior: 'caret-end', pendingTypeSeed: null }
  }

  const hasPrimaryModifier = event.ctrlKey || event.metaKey
  const normalizedKey = event.key.toLowerCase()

  if (isCurrentRegionSelectionShortcut(event)) {
    const currentRegion = navigation?.resolveCurrentRegion(activeCell) ?? null
    if (currentRegion) {
      return { kind: 'select-range', cell: activeCell, range: currentRegion }
    }
    return { kind: 'handled' }
  }

  if (hasPrimaryModifier && !event.altKey && normalizedKey === 'a') {
    const currentRegion = navigation?.resolveCurrentRegion(activeCell) ?? null
    if (currentRegion && !rectanglesEqual(currentRegion, currentSelectionRange)) {
      return { kind: 'select-range', cell: activeCell, range: currentRegion }
    }
    return { kind: 'select-all' }
  }

  if (!event.altKey && event.key === ' ' && hasPrimaryModifier && event.shiftKey) {
    return { kind: 'select-all' }
  }

  if (!event.altKey && event.key === ' ' && hasPrimaryModifier) {
    return { kind: 'select-column', col: activeCell[0], row: activeCell[1] }
  }

  if (!event.altKey && event.key === ' ' && event.shiftKey) {
    return { kind: 'select-row', col: activeCell[0], row: activeCell[1] }
  }

  if (event.key === 'Home' && !event.altKey) {
    const nextCell = hasPrimaryModifier ? ([0, 0] as Item) : ([0, activeCell[1]] as Item)
    if (event.shiftKey) {
      return {
        kind: 'extend-selection',
        anchor: anchorCell,
        target: nextCell,
      }
    }
    return { kind: 'move-selection', cell: nextCell }
  }

  if (event.key === 'End' && !event.altKey) {
    const nextCell = hasPrimaryModifier ? ([MAX_COLS - 1, MAX_ROWS - 1] as Item) : ([MAX_COLS - 1, activeCell[1]] as Item)
    if (event.shiftKey) {
      return {
        kind: 'extend-selection',
        anchor: anchorCell,
        target: nextCell,
      }
    }
    return { kind: 'move-selection', cell: nextCell }
  }

  if ((event.key === 'PageUp' || event.key === 'PageDown') && !event.altKey) {
    const jumpRows = resolvePageJumpRows(pageJumpRows)
    const nextCell = clampCell([activeCell[0], activeCell[1] + (event.key === 'PageDown' ? jumpRows : -jumpRows)])
    if (event.shiftKey) {
      return {
        kind: 'extend-selection',
        anchor: anchorCell,
        target: nextCell,
      }
    }
    return { kind: 'move-selection', cell: nextCell }
  }

  if (event.key === 'Enter' && !event.altKey && !event.ctrlKey && !event.metaKey) {
    if (isMultiCellRange(currentSelectionRange)) {
      return {
        kind: 'move-selection-in-range',
        cell: moveWithinRange(
          resolveRangeNavigationCell(selectedCell, currentSelectionCell, currentSelectionRange),
          currentSelectionRange,
          'vertical',
          Boolean(event.shiftKey),
        ),
        range: currentSelectionRange,
      }
    }
    return {
      kind: 'move-selection',
      cell: clampCell([activeCell[0], activeCell[1] + (event.shiftKey ? -1 : 1)]),
    }
  }

  if (event.key === 'Tab' && !event.altKey && !event.ctrlKey && !event.metaKey) {
    if (isMultiCellRange(currentSelectionRange)) {
      return {
        kind: 'move-selection-in-range',
        cell: moveWithinRange(
          resolveRangeNavigationCell(selectedCell, currentSelectionCell, currentSelectionRange),
          currentSelectionRange,
          'horizontal',
          Boolean(event.shiftKey),
        ),
        range: currentSelectionRange,
      }
    }
    return {
      kind: 'move-selection',
      cell: clampCell([activeCell[0] + (event.shiftKey ? -1 : 1), activeCell[1]]),
    }
  }

  if ((event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'ArrowLeft' || event.key === 'ArrowRight') && !event.altKey) {
    const delta: Item =
      event.key === 'ArrowUp' ? [0, -1] : event.key === 'ArrowDown' ? [0, 1] : event.key === 'ArrowLeft' ? [-1, 0] : [1, 0]
    const direction: GridNavigationDirection =
      event.key === 'ArrowUp' ? 'up' : event.key === 'ArrowDown' ? 'down' : event.key === 'ArrowLeft' ? 'left' : 'right'
    const nextCell = hasPrimaryModifier
      ? (navigation?.resolveDataEdge(activeCell, direction) ?? moveSelectionToEdge(activeCell, direction))
      : clampCell([activeCell[0] + delta[0], activeCell[1] + delta[1]])

    if (event.shiftKey) {
      return {
        kind: 'extend-selection',
        anchor: anchorCell,
        target: nextCell,
      }
    }

    return { kind: 'move-selection', cell: nextCell }
  }

  if (isClearCellKey(event)) {
    return { kind: 'clear-cell', pendingTypeSeed: null }
  }

  if (isScrollActiveCellShortcut(event)) {
    return { kind: 'scroll-active-cell' }
  }

  if (isStructuralDeleteShortcut(event)) {
    if (selectedRowRanges && selectedRowRanges.length > 0 && (!selectedColumnRanges || selectedColumnRanges.length === 0)) {
      return { kind: 'delete-selected-rows', ranges: selectedRowRanges }
    }
    if (selectedColumnRanges && selectedColumnRanges.length > 0 && (!selectedRowRanges || selectedRowRanges.length === 0)) {
      return { kind: 'delete-selected-columns', ranges: selectedColumnRanges }
    }
    return { kind: 'handled' }
  }

  if (isFillSelectionShortcut(event)) {
    if (!currentSelectionRange || (currentSelectionRange.width === 1 && currentSelectionRange.height === 1)) {
      return { kind: 'handled' }
    }
    return {
      kind: 'fill-range',
      source: {
        x: selectedActionCell[0],
        y: selectedActionCell[1],
        width: 1,
        height: 1,
      },
      target: currentSelectionRange,
    }
  }

  if (isFillShortcut(event)) {
    if (!currentSelectionRange) {
      return { kind: 'handled' }
    }
    if (normalizedKey === 'd') {
      if (currentSelectionRange.height <= 1) {
        return { kind: 'handled' }
      }
      return {
        kind: 'fill-range',
        source: {
          x: currentSelectionRange.x,
          y: currentSelectionRange.y,
          width: currentSelectionRange.width,
          height: 1,
        },
        target: {
          x: currentSelectionRange.x,
          y: currentSelectionRange.y + 1,
          width: currentSelectionRange.width,
          height: currentSelectionRange.height - 1,
        },
      }
    }
    if (currentSelectionRange.width <= 1) {
      return { kind: 'handled' }
    }
    return {
      kind: 'fill-range',
      source: {
        x: currentSelectionRange.x,
        y: currentSelectionRange.y,
        width: 1,
        height: currentSelectionRange.height,
      },
      target: {
        x: currentSelectionRange.x + 1,
        y: currentSelectionRange.y,
        width: currentSelectionRange.width - 1,
        height: currentSelectionRange.height,
      },
    }
  }

  if (hasPrimaryModifier && !event.altKey) {
    if (normalizedKey === 'c') {
      return { kind: 'clipboard-copy' }
    }
    if (normalizedKey === 'x') {
      return { kind: 'clipboard-cut' }
    }
    if (normalizedKey === 'v') {
      return { kind: 'clipboard-paste', target: activeCell, valuesOnly: event.shiftKey === true }
    }
  }

  if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    const seed = `${pendingTypeSeed ?? ''}${event.key}`
    return {
      kind: 'begin-edit',
      seed,
      selectionBehavior: 'caret-end',
      pendingTypeSeed: seed,
    }
  }

  return { kind: 'none' }
}
