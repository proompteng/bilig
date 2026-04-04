import { describe, expect, test } from "vitest";
import {
  resolveFillHandleOverlayBounds,
  resolveFillHandlePreviewRange,
} from "../gridFillHandle.js";

describe("gridFillHandle", () => {
  test("resolves preview ranges by extending the dominant drag axis", () => {
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [6, 4])).toEqual({
      x: 2,
      y: 3,
      width: 5,
      height: 2,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [1, 4])).toEqual({
      x: 1,
      y: 3,
      width: 3,
      height: 2,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 8])).toEqual({
      x: 2,
      y: 3,
      width: 2,
      height: 6,
    });

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 1])).toEqual({
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
