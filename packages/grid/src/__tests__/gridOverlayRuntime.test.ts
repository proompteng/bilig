import { describe, expect, it } from 'vitest'
import { OVERLAY_INSTANCE_FLOATS_V3, packOverlayBatchV3 } from '../renderer-v3/overlay-layer.js'
import { GridOverlayRuntime } from '../runtime/gridOverlayRuntime.js'

describe('renderer-v3 overlay runtime', () => {
  it('packs dynamic overlays into a compact instance buffer', () => {
    const batch = packOverlayBatchV3({
      axisSeqX: 3,
      axisSeqY: 4,
      cameraSeq: 2,
      instances: [
        {
          alpha: 0.5,
          color: '#336699',
          height: 20,
          kind: 'selection',
          width: 100,
          x: 10,
          y: 12,
          z: 1,
        },
      ],
      seq: 1,
    })

    expect(batch.count).toBe(1)
    expect(batch.kindMask).toBe(1)
    expect(batch.instances).toHaveLength(OVERLAY_INSTANCE_FLOATS_V3)
    expect(Array.from(batch.instances.slice(0, 5))).toEqual([0, 10, 12, 100, 20])
  })

  it('updates overlay batches without touching data tile state', () => {
    const runtime = new GridOverlayRuntime()

    expect(runtime.set('active', { color: '#1a73e8', height: 22, kind: 'activeCell', width: 104, x: 0, y: 0 })).toBe(1)
    expect(runtime.set('resize', { color: '#202124', height: 400, kind: 'resizeGuide', width: 1, x: 300, y: 0 })).toBe(2)

    const batch = runtime.buildBatch({ axisSeqX: 5, axisSeqY: 6, cameraSeq: 7 })

    expect(batch).toMatchObject({ axisSeqX: 5, axisSeqY: 6, cameraSeq: 7, count: 2, seq: 2 })
    expect(batch.kindMask).toBe((1 << 1) | (1 << 3))

    runtime.clearKind('resizeGuide')
    expect(runtime.buildBatch({ axisSeqX: 5, axisSeqY: 6, cameraSeq: 8 })).toMatchObject({ count: 1, seq: 3 })
  })
})
