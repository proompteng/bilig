import type { Item, Rectangle } from "@glideapps/glide-data-grid";

export interface FillHandleOverlayBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export function resolveFillHandlePreviewRange(
  sourceRange: Rectangle,
  pointerCell: Item,
): Rectangle | null {
  const sourceLeft = sourceRange.x;
  const sourceTop = sourceRange.y;
  const sourceRight = sourceRange.x + sourceRange.width - 1;
  const sourceBottom = sourceRange.y + sourceRange.height - 1;

  const leftDelta = pointerCell[0] < sourceLeft ? sourceLeft - pointerCell[0] : 0;
  const rightDelta = pointerCell[0] > sourceRight ? pointerCell[0] - sourceRight : 0;
  const upDelta = pointerCell[1] < sourceTop ? sourceTop - pointerCell[1] : 0;
  const downDelta = pointerCell[1] > sourceBottom ? pointerCell[1] - sourceBottom : 0;

  const horizontalDelta = Math.max(leftDelta, rightDelta);
  const verticalDelta = Math.max(upDelta, downDelta);
  if (horizontalDelta === 0 && verticalDelta === 0) {
    return null;
  }

  if (horizontalDelta >= verticalDelta) {
    if (rightDelta > 0) {
      return {
        x: sourceLeft,
        y: sourceTop,
        width: pointerCell[0] - sourceLeft + 1,
        height: sourceRange.height,
      };
    }
    return {
      x: pointerCell[0],
      y: sourceTop,
      width: sourceRight - pointerCell[0] + 1,
      height: sourceRange.height,
    };
  }

  if (downDelta > 0) {
    return {
      x: sourceLeft,
      y: sourceTop,
      width: sourceRange.width,
      height: pointerCell[1] - sourceTop + 1,
    };
  }
  return {
    x: sourceLeft,
    y: pointerCell[1],
    width: sourceRange.width,
    height: sourceBottom - pointerCell[1] + 1,
  };
}

export function resolveFillHandleOverlayBounds(options: {
  sourceRange: Rectangle;
  getCellBounds: (col: number, row: number) => Rectangle | undefined;
  hostBounds: Pick<DOMRect, "left" | "top">;
  size?: number;
}): FillHandleOverlayBounds | undefined {
  const { getCellBounds, hostBounds, size = 12, sourceRange } = options;
  const anchorBounds = getCellBounds(
    sourceRange.x + sourceRange.width - 1,
    sourceRange.y + sourceRange.height - 1,
  );
  if (!anchorBounds) {
    return undefined;
  }

  return {
    x: anchorBounds.x - hostBounds.left + anchorBounds.width - size / 2,
    y: anchorBounds.y - hostBounds.top + anchorBounds.height - size / 2,
    width: size,
    height: size,
  };
}
