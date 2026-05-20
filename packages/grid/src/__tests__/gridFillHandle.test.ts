import { describe, expect, test } from 'vitest'
import {
  resolveFillHandleHitTargetBounds,
  resolveFillHandlePreviewBounds,
  resolveFillHandlePreviewRange,
  resolveFillHandleSelectionRange,
} from '../gridFillHandle.js'

describe('gridFillHandle', () => {
  test('resolves preview ranges to only the fill target cells on the dominant drag axis', () => {
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [6, 4])).toEqual({
      x: 4,
      y: 3,
      width: 3,
      height: 2,
    })

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [1, 4])).toEqual({
      x: 1,
      y: 3,
      width: 1,
      height: 2,
    })

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 8])).toEqual({
      x: 2,
      y: 5,
      width: 2,
      height: 4,
    })

    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 1])).toEqual({
      x: 2,
      y: 1,
      width: 2,
      height: 2,
    })
  })

  test('resolves the post-fill selection to include the source and target ranges', () => {
    expect(resolveFillHandleSelectionRange({ x: 2, y: 3, width: 2, height: 2 }, { x: 4, y: 3, width: 3, height: 2 })).toEqual({
      x: 2,
      y: 3,
      width: 5,
      height: 2,
    })

    expect(resolveFillHandleSelectionRange({ x: 2, y: 3, width: 2, height: 2 }, { x: 2, y: 1, width: 2, height: 2 })).toEqual({
      x: 2,
      y: 1,
      width: 2,
      height: 4,
    })
  })

  test('returns null when the pointer stays inside the source range', () => {
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [2, 3])).toBeNull()
    expect(resolveFillHandlePreviewRange({ x: 2, y: 3, width: 2, height: 2 }, [3, 4])).toBeNull()
  })

  test('computes preview bounds from the visible portion of the target range', () => {
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
    })
  })

  test('computes the hit target from the visible handle center', () => {
    expect(
      resolveFillHandleHitTargetBounds({
        hostBounds: { width: 400, height: 300 },
        visualBounds: { x: 226.5, y: 140.5, width: 7, height: 7 },
      }),
    ).toEqual({
      x: 225,
      y: 139,
      width: 10,
      height: 10,
    })
  })

  test('hides the hit target when the visible handle is outside the grid body', () => {
    expect(
      resolveFillHandleHitTargetBounds({
        hostBounds: { width: 400, height: 300 },
        minX: 46,
        minY: 24,
        visualBounds: { x: 20, y: -20, width: 7, height: 7 },
      }),
    ).toBeUndefined()
  })

  test('clips the hit target to the visible grid body instead of intercepting headers', () => {
    expect(
      resolveFillHandleHitTargetBounds({
        hostBounds: { width: 400, height: 300 },
        minX: 46,
        minY: 24,
        visualBounds: { x: 43.5, y: 22.5, width: 7, height: 7 },
      }),
    ).toEqual({
      x: 46,
      y: 24,
      width: 6,
      height: 7,
    })
  })

  test('clips the hit target at the viewport edge instead of covering footer chrome', () => {
    expect(
      resolveFillHandleHitTargetBounds({
        hostBounds: { width: 320, height: 100 },
        minX: 46,
        minY: 24,
        visualBounds: { x: 312.5, y: 92.5, width: 7, height: 7 },
      }),
    ).toEqual({
      x: 311,
      y: 91,
      width: 9,
      height: 9,
    })
  })
})
