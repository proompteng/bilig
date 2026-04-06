import { describe, expect, test } from "vitest";
import {
  resolveFillHandleOverlayBounds,
  resolveFillHandlePreviewBounds,
  resolveFillHandlePreviewRange,
  resolveFillHandleSelectionRange,
} from "../gridFillHandle.js";

describe("gridFillHandle", () => {
  test("resolves preview ranges to only the fill target cells on the dominant drag axis", () => {
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [6, 4])).toEqual({
      x: 4,
      y: 3,
      width: 3,
      height: 2,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [1, 4])).toEqual({
      x: 1,
      y: 3,
      width: 1,
      height: 2,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 8])).toEqual({
      x: 2,
      y: 5,
      width: 2,
      height: 4,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 1])).toEqual({
      x: 2,
      y: 1,
      width: 2,
      height: 2,
    });
  });

  test("resolves the post-fill selection to include the source and target ranges", () => {
    expect(
      resolveFillHandleSelectionRange(
        { x: 2, y: 3, width: 2, height: 2 },
        { x: 4, y: 3, width: 3, height: 2 },
      ),
    ).toEqual({
      x: 2,
      y: 3,
      width: 5,
      height: 2,
    });

    expect(
      resolveFillHandleSelectionRange(
        { x: 2, y: 3, width: 2, height: 2 },
        { x: 2, y: 1, width: 2, height: 2 },
      ),
    ).toEqual({
      x: 2,
      y: 1,
      width: 2,
      height: 4,
    });
  });

  test("returns null when the pointer stays inside the source range", () => {
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [2, 3])).toBeNull();
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 4])).toBeNull();
  });

  test("computes preview bounds from the visible portion of the target range", () => {
    expect(
      resolveFillHandlePreviewBounds({
        previewRange: { x: 4, y: 3, width: 3, height: 2 },
        visibleRange: { x: 3, y: 2, width: 3, height: 3 },
        hostBounds: { left: 100, top: 200 },
        getCellBounds: (col, row) => ({
          x: 100 + col * 80,
          y: 200 + row * 24,
          width: 80,
          height: 24,
        }),
      }),
    ).toEqual({
      x: 320,
      y: 72,
      width: 160,
      height: 48,
    });
  });

  test("computes overlay bounds from the trailing cell corner", () => {
    expect(
      resolveFillHandleOverlayBounds({
        sourceRange: { x: 1, y: 2, width: 2, height: 3 },
        hostBounds: { left: 100, top: 200 },
        getCellBounds: (col, row) =>
          col === 2 && row === 4 ? { x: 250, y: 320, width: 80, height: 24 } : undefined,
      }),
    ).toEqual({
      x: 224,
      y: 138,
      width: 12,
      height: 12,
    });
  });
});
