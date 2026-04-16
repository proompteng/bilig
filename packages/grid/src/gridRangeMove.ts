import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { Item, Rectangle } from './gridTypes.js'

const RANGE_MOVE_BORDER_THRESHOLD = 6

export function resolveSelectionBounds(
  sourceRange: Rectangle,
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
): Rectangle | undefined {
  const startBounds = getCellBounds(sourceRange.x, sourceRange.y)
  const endBounds = getCellBounds(sourceRange.x + sourceRange.width - 1, sourceRange.y + sourceRange.height - 1)
  if (!startBounds || !endBounds) {
    return undefined
  }
  return {
    x: startBounds.x,
    y: startBounds.y,
    width: endBounds.x + endBounds.width - startBounds.x,
    height: endBounds.y + endBounds.height - startBounds.y,
  }
}

export function resolveSelectionMoveAnchorCell(
  clientX: number,
  clientY: number,
  sourceRange: Rectangle | null | undefined,
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
  threshold = RANGE_MOVE_BORDER_THRESHOLD,
): Item | null {
  if (!sourceRange) {
    return null
  }
  const selectionBounds = resolveSelectionBounds(sourceRange, getCellBounds)
  if (
    !selectionBounds ||
    clientX < selectionBounds.x ||
    clientX >= selectionBounds.x + selectionBounds.width ||
    clientY < selectionBounds.y ||
    clientY >= selectionBounds.y + selectionBounds.height
  ) {
    return null
  }

  let pointerCell: Item | null = null
  const sourceRight = sourceRange.x + sourceRange.width - 1
  const sourceBottom = sourceRange.y + sourceRange.height - 1
  for (let row = sourceRange.y; row <= sourceBottom && pointerCell === null; row += 1) {
    for (let col = sourceRange.x; col <= sourceRight; col += 1) {
      const cellBounds = getCellBounds(col, row)
      if (
        cellBounds &&
        clientX >= cellBounds.x &&
        clientX < cellBounds.x + cellBounds.width &&
        clientY >= cellBounds.y &&
        clientY < cellBounds.y + cellBounds.height
      ) {
        pointerCell = [col, row]
        break
      }
    }
  }

  if (!pointerCell) {
    return null
  }
  const cellBounds = getCellBounds(pointerCell[0], pointerCell[1])
  if (!cellBounds) {
    return null
  }
  if (
    clientX < cellBounds.x ||
    clientX >= cellBounds.x + cellBounds.width ||
    clientY < cellBounds.y ||
    clientY >= cellBounds.y + cellBounds.height
  ) {
    return null
  }

  const localX = clientX - cellBounds.x
  const localY = clientY - cellBounds.y
  return (pointerCell[0] === sourceRange.x && localX < threshold) ||
    (pointerCell[0] === sourceRight && localX >= cellBounds.width - threshold) ||
    (pointerCell[1] === sourceRange.y && localY < threshold) ||
    (pointerCell[1] === sourceBottom && localY >= cellBounds.height - threshold)
    ? pointerCell
    : null
}

export function isSelectionMoveHandleHit(
  clientX: number,
  clientY: number,
  sourceRange: Rectangle | null | undefined,
  getCellBounds: (col: number, row: number) => Rectangle | undefined,
  threshold = RANGE_MOVE_BORDER_THRESHOLD,
): boolean {
  return resolveSelectionMoveAnchorCell(clientX, clientY, sourceRange, getCellBounds, threshold) !== null
}

export function resolveMovedRange(sourceRange: Rectangle, pointerCell: Item, anchorOffset: Item): Rectangle {
  const maxX = Math.max(0, MAX_COLS - sourceRange.width)
  const maxY = Math.max(0, MAX_ROWS - sourceRange.height)
  const nextX = Math.min(maxX, Math.max(0, pointerCell[0] - anchorOffset[0]))
  const nextY = Math.min(maxY, Math.max(0, pointerCell[1] - anchorOffset[1]))
  return {
    x: nextX,
    y: nextY,
    width: sourceRange.width,
    height: sourceRange.height,
  }
}

export function sameRectangle(left: Rectangle | null | undefined, right: Rectangle | null | undefined): boolean {
  return (
    left === right ||
    (left !== null &&
      left !== undefined &&
      right !== null &&
      right !== undefined &&
      left.x === right.x &&
      left.y === right.y &&
      left.width === right.width &&
      left.height === right.height)
  )
}
