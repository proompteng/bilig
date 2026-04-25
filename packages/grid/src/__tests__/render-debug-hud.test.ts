import { describe, expect, test } from 'vitest'
import { EMPTY_GRID_GPU_COUNTERS } from '../renderer/grid-render-counters.js'
import { formatRenderDebugHud, isRenderDebugSnapshotInsideBudget } from '../renderer-v2/render-debug-hud.js'

describe('render-debug-hud', () => {
  test('formats useful render counters', () => {
    expect(
      formatRenderDebugHud({
        frameMs: 8.333,
        gpu: { ...EMPTY_GRID_GPU_COUNTERS, drawCalls: 4, submitCount: 1, vertexUploadBytes: 2048 },
        inputToDrawMs: 6.25,
      }),
    ).toContain('uploads 2.0KB')
  })

  test('checks render budgets', () => {
    expect(
      isRenderDebugSnapshotInsideBudget({
        maxBufferAllocations: 0,
        maxFrameMs: 16.7,
        maxInputToDrawMs: 12,
        maxTileMisses: 0,
        maxVertexUploadBytes: 0,
        snapshot: {
          frameMs: 8,
          gpu: EMPTY_GRID_GPU_COUNTERS,
          inputToDrawMs: 7,
        },
      }),
    ).toBe(true)
  })
})
