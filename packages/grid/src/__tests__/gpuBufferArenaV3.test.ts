import { describe, expect, it } from 'vitest'
import { GpuBufferArenaV3 } from '../renderer-v3/gpu-buffer-arena.js'

describe('GpuBufferArenaV3', () => {
  it('reuses released buffers by layout and capacity class', () => {
    let nextId = 1
    const arena = new GpuBufferArenaV3(({ capacityBytes, layout }) => ({ capacityBytes, id: nextId++, layout }))

    const first = arena.acquire('rectInstances', 300)
    expect(first.capacityBytes).toBe(512)
    expect(first.usedBytes).toBe(300)

    arena.release(first)
    const second = arena.acquire('rectInstances', 280)

    expect(second.buffer).toBe(first.buffer)
    expect(second.usedBytes).toBe(280)
    expect(arena.stats()).toMatchObject({ creates: 1, reuses: 1 })
  })

  it('keeps layout classes isolated and trims free buffers on demand', () => {
    const destroyed: unknown[] = []
    const arena = new GpuBufferArenaV3(
      ({ capacityBytes, layout }) => ({ capacityBytes, layout }),
      (buffer) => destroyed.push(buffer),
    )

    const rect = arena.acquire('rectInstances', 512)
    const overlay = arena.acquire('overlayInstances', 512)
    arena.release(rect)
    arena.release(overlay)

    expect(arena.acquire('textRuns', 512).buffer).not.toBe(rect.buffer)
    expect(arena.trim(512)).toBeGreaterThanOrEqual(512)
    expect(destroyed).toHaveLength(1)
    expect(arena.stats().destroys).toBe(1)
  })
})
