import type { Item, Rectangle } from './gridTypes.js'

export const GRID_FILL_HANDLE_VISUAL_SIZE = 7
export const GRID_FILL_HANDLE_HIT_SIZE = 10

export interface FillHandleOverlayBounds {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}

export function resolveFillHandlePreviewRange(sourceRange: Rectangle, pointerCell: Item): Rectangle | null {
  const sourceLeft = sourceRange.x
  const sourceTop = sourceRange.y
  const sourceRight = sourceRange.x + sourceRange.width - 1
  const sourceBottom = sourceRange.y + sourceRange.height - 1

  const leftDelta = pointerCell[0] < sourceLeft ? sourceLeft - pointerCell[0] : 0
  const rightDelta = pointerCell[0] > sourceRight ? pointerCell[0] - sourceRight : 0
  const upDelta = pointerCell[1] < sourceTop ? sourceTop - pointerCell[1] : 0
  const downDelta = pointerCell[1] > sourceBottom ? pointerCell[1] - sourceBottom : 0

  const horizontalDelta = Math.max(leftDelta, rightDelta)
  const verticalDelta = Math.max(upDelta, downDelta)
  if (horizontalDelta === 0 && verticalDelta === 0) {
    return null
  }

  if (horizontalDelta >= verticalDelta) {
    if (rightDelta > 0) {
      return {
        x: sourceRight + 1,
        y: sourceTop,
        width: rightDelta,
        height: sourceRange.height,
      }
    }
    return {
      x: pointerCell[0],
      y: sourceTop,
      width: leftDelta,
      height: sourceRange.height,
    }
  }

  if (downDelta > 0) {
    return {
      x: sourceLeft,
      y: sourceBottom + 1,
      width: sourceRange.width,
      height: downDelta,
    }
  }
  return {
    x: sourceLeft,
    y: pointerCell[1],
    width: sourceRange.width,
    height: upDelta,
  }
}

export function resolveFillHandleSelectionRange(sourceRange: Rectangle, previewRange: Rectangle): Rectangle {
  const sourceRight = sourceRange.x + sourceRange.width - 1
  const sourceBottom = sourceRange.y + sourceRange.height - 1
  const previewRight = previewRange.x + previewRange.width - 1
  const previewBottom = previewRange.y + previewRange.height - 1

  const left = Math.min(sourceRange.x, previewRange.x)
  const top = Math.min(sourceRange.y, previewRange.y)
  const right = Math.max(sourceRight, previewRight)
  const bottom = Math.max(sourceBottom, previewBottom)

  return {
    x: left,
    y: top,
    width: right - left + 1,
    height: bottom - top + 1,
  }
}

export function resolveFillHandlePreviewBounds(options: {
  previewRange: Rectangle
  visibleRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  getCellBounds: (col: number, row: number) => Rectangle | undefined
  hostBounds: Pick<DOMRect, 'left' | 'top'>
}): Rectangle | undefined {
  const { getCellBounds, hostBounds, previewRange, visibleRange } = options
  const startCol = Math.max(previewRange.x, visibleRange.x)
  const startRow = Math.max(previewRange.y, visibleRange.y)
  const endCol = Math.min(previewRange.x + previewRange.width - 1, visibleRange.x + visibleRange.width - 1)
  const endRow = Math.min(previewRange.y + previewRange.height - 1, visibleRange.y + visibleRange.height - 1)
  if (startCol > endCol || startRow > endRow) {
    return undefined
  }

  const startBounds = getCellBounds(startCol, startRow)
  const endBounds = getCellBounds(endCol, endRow)
  if (!startBounds || !endBounds) {
    return undefined
  }

  return {
    x: startBounds.x - hostBounds.left,
    y: startBounds.y - hostBounds.top,
    width: endBounds.x + endBounds.width - startBounds.x,
    height: endBounds.y + endBounds.height - startBounds.y,
  }
}

export function resolveFillHandleHitTargetBounds(options: {
  visualBounds: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
  hostBounds: Pick<DOMRect, 'width' | 'height'>
  minX?: number
  minY?: number
  size?: number
}): FillHandleOverlayBounds | undefined {
  const { hostBounds, minX = 0, minY = 0, size = GRID_FILL_HANDLE_HIT_SIZE, visualBounds } = options
  const centerX = visualBounds.x + visualBounds.width / 2
  const centerY = visualBounds.y + visualBounds.height / 2

  const resolvedBounds = {
    x: centerX - size / 2,
    y: centerY - size / 2,
    width: size,
    height: size,
  }

  const clippedLeft = Math.max(minX, resolvedBounds.x)
  const clippedTop = Math.max(minY, resolvedBounds.y)
  const clippedRight = Math.min(hostBounds.width, resolvedBounds.x + resolvedBounds.width)
  const clippedBottom = Math.min(hostBounds.height, resolvedBounds.y + resolvedBounds.height)

  if (clippedRight <= clippedLeft || clippedBottom <= clippedTop) {
    return undefined
  }

  return {
    x: clippedLeft,
    y: clippedTop,
    width: clippedRight - clippedLeft,
    height: clippedBottom - clippedTop,
  }
}
