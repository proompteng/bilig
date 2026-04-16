import { formatAddress } from '@bilig/formula'
import { createColumnSliceSelection, createGridSelection, createRangeSelection, createRowSliceSelection } from './gridSelection.js'
import type { HeaderSelection } from './gridPointer.js'
import type { GridSelection, Item } from './gridTypes.js'

export function resolveHeaderDragSelection(
  headerAnchor: HeaderSelection,
  targetIndex: number,
  selectedCell: Item,
): {
  selection: GridSelection
  addr: string
} {
  if (headerAnchor.kind === 'column') {
    return {
      selection: createColumnSliceSelection(headerAnchor.index, targetIndex, selectedCell[1]),
      addr: formatAddress(selectedCell[1], headerAnchor.index),
    }
  }

  return {
    selection: createRowSliceSelection(selectedCell[0], headerAnchor.index, targetIndex),
    addr: formatAddress(headerAnchor.index, selectedCell[0]),
  }
}

export function resolveBodyDragSelection(anchorCell: Item, pointerCell: Item): GridSelection {
  return createRangeSelection(createGridSelection(anchorCell[0], anchorCell[1]), anchorCell, pointerCell)
}

export function resolveBodyPointerUpResult(
  anchorCell: Item,
  pointerCell: Item,
  didMove: boolean,
): {
  selection: GridSelection | null
  addr: string | null
  clickedCell: Item | null
  shouldSetDragExpiry: boolean
} {
  if (!didMove) {
    return {
      selection: null,
      addr: null,
      clickedCell: anchorCell,
      shouldSetDragExpiry: false,
    }
  }

  return {
    selection: resolveBodyDragSelection(anchorCell, pointerCell),
    addr: formatAddress(anchorCell[1], anchorCell[0]),
    clickedCell: null,
    shouldSetDragExpiry: true,
  }
}
