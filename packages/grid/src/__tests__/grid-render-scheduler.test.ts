import { describe, expect, test } from 'vitest'
import { GridRenderScheduler } from '../renderer/grid-render-scheduler.js'

describe('GridRenderScheduler', () => {
  test('coalesces draw requests into one frame', () => {
    const callbacks: FrameRequestCallback[] = []
    const scheduler = new GridRenderScheduler((callback) => {
      callbacks.push(callback)
      return callbacks.length
    })
    let draws = 0

    scheduler.requestDraw(() => {
      draws += 1
    })
    scheduler.requestDraw(() => {
      draws += 1
    })

    expect(callbacks).toHaveLength(1)
    callbacks[0]?.(10)
    expect(draws).toBe(1)
  })
})
