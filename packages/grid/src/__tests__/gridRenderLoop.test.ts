import { describe, expect, test, vi } from 'vitest'
import { GridRenderLoop } from '../renderer-v2/gridRenderLoop.js'

describe('GridRenderLoop', () => {
  test('coalesces requests and draws the latest callback', () => {
    let callback: FrameRequestCallback | null = null
    const requestFrame = vi.fn((next: FrameRequestCallback) => {
      callback = next
      return 1
    })
    const loop = new GridRenderLoop(requestFrame, vi.fn())
    const first = vi.fn()
    const second = vi.fn()

    loop.requestDraw(first)
    loop.requestDraw(second)
    callback?.(100)

    expect(requestFrame).toHaveBeenCalledTimes(1)
    expect(first).not.toHaveBeenCalled()
    expect(second).toHaveBeenCalledTimes(1)
  })

  test('cancels pending draws', () => {
    const cancelFrame = vi.fn()
    const loop = new GridRenderLoop(() => 7, cancelFrame)
    const draw = vi.fn()

    loop.requestDraw(draw)
    loop.cancel()

    expect(cancelFrame).toHaveBeenCalledWith(7)
  })
})
