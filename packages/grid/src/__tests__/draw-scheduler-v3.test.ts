import { describe, expect, test, vi } from 'vitest'
import {
  GridDrawSchedulerV3,
  TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS,
  TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS,
  shouldDeferTypeGpuV3PreloadSync,
} from '../renderer-v3/draw-scheduler.js'
import { GridRenderLoop } from '../renderer-v3/gridRenderLoop.js'

describe('GridDrawSchedulerV3', () => {
  test('defers preload sync only while input or camera movement is fresh', () => {
    expect(
      shouldDeferTypeGpuV3PreloadSync({
        camera: null,
        lastScrollSignalAt: 1_000,
        now: 1_000 + TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS - 1,
      }),
    ).toBe(true)

    expect(
      shouldDeferTypeGpuV3PreloadSync({
        camera: {
          updatedAt: 2_000,
          velocityX: 0,
          velocityY: 0,
        },
        lastScrollSignalAt: 1_000,
        now: 2_000 + TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS + 1,
      }),
    ).toBe(false)
  })

  test('schedules an idle preload draw when the input lane defers warm tile work', () => {
    const draw = vi.fn()
    const scheduled: Array<() => void> = []
    const frames: FrameRequestCallback[] = []
    const renderLoop = new GridRenderLoop((callback) => {
      frames.push(callback)
      return frames.length
    }, vi.fn())
    const scheduler = new GridDrawSchedulerV3(
      (callback, delay) => {
        expect(delay).toBe(TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS)
        scheduled.push(callback)
        return scheduled.length
      },
      vi.fn(),
      () => 1_000,
      renderLoop,
    )

    scheduler.noteInputSignal(1_000)
    expect(
      scheduler.resolveFrame({
        camera: null,
        requestIdlePreloadDraw: () => scheduler.requestDraw(draw),
      }),
    ).toEqual({
      deferPreloadSync: true,
      syncPreloadPanes: false,
    })

    scheduled[0]?.()
    expect(frames).toHaveLength(1)
    frames[0]?.(1_001)
    expect(draw).toHaveBeenCalledTimes(1)
  })
})
