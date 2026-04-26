import { describe, expect, it, vi } from 'vitest'
import { WorkbookPaneBufferCache } from '../renderer-v2/pane-buffer-cache.js'
import type { RectInstanceVertexBuffer, TextInstanceVertexBuffer } from '../renderer-v2/typegpu-backend.js'

function rectBuffer(destroy: () => void): RectInstanceVertexBuffer {
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion
  return { destroy } as unknown as RectInstanceVertexBuffer
}

function textBuffer(destroy: () => void): TextInstanceVertexBuffer {
  // oxlint-disable-next-line typescript-eslint/no-unsafe-type-assertion
  return { destroy } as unknown as TextInstanceVertexBuffer
}

describe('WorkbookPaneBufferCache', () => {
  it('releases pruned pane vertex buffers into reusable free lists', () => {
    const cache = new WorkbookPaneBufferCache()
    const entry = cache.get('body')
    const rectDestroy = vi.fn()
    const textDestroy = vi.fn()
    const rect = rectBuffer(rectDestroy)
    const text = textBuffer(textDestroy)
    entry.rectBuffer = rect
    entry.rectCapacity = 128
    entry.textBuffer = text
    entry.textCapacity = 64

    cache.delete('body')

    expect(cache.acquireRectBuffer(64)).toEqual({ buffer: rect, capacity: 128 })
    expect(cache.acquireTextBuffer(64)).toEqual({ buffer: text, capacity: 64 })
    expect(rectDestroy).not.toHaveBeenCalled()
    expect(textDestroy).not.toHaveBeenCalled()
  })

  it('chooses the smallest reusable buffer that satisfies the request', () => {
    const cache = new WorkbookPaneBufferCache()
    const small = rectBuffer(vi.fn())
    const large = rectBuffer(vi.fn())
    cache.releaseRectBuffer(large, 512)
    cache.releaseRectBuffer(small, 128)

    expect(cache.acquireRectBuffer(100)).toEqual({ buffer: small, capacity: 128 })
    expect(cache.acquireRectBuffer(100)).toEqual({ buffer: large, capacity: 512 })
  })

  it('destroys retained free-list buffers on dispose', () => {
    const cache = new WorkbookPaneBufferCache()
    const rectDestroy = vi.fn()
    const textDestroy = vi.fn()
    const rect = rectBuffer(rectDestroy)
    const text = textBuffer(textDestroy)
    cache.releaseRectBuffer(rect, 128)
    cache.releaseTextBuffer(text, 64)

    cache.dispose()

    expect(rectDestroy).toHaveBeenCalledTimes(1)
    expect(textDestroy).toHaveBeenCalledTimes(1)
  })
})
