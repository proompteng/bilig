export type OverlayKindV3 = 'selection' | 'activeCell' | 'fillHandle' | 'resizeGuide' | 'hover' | 'remoteCursor' | 'frozenSeparator'

export interface OverlayInstanceV3 {
  readonly kind: OverlayKindV3
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: string
  readonly alpha?: number | undefined
  readonly z?: number | undefined
}

export interface OverlayBatchV3 {
  readonly seq: number
  readonly cameraSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly instances: Float32Array
  readonly count: number
  readonly kindMask: number
}

export const OVERLAY_INSTANCE_FLOATS_V3 = 10

const KIND_TAGS: Record<OverlayKindV3, number> = {
  selection: 0,
  activeCell: 1,
  fillHandle: 2,
  resizeGuide: 3,
  hover: 4,
  remoteCursor: 5,
  frozenSeparator: 6,
}

export function packOverlayBatchV3(input: {
  readonly seq: number
  readonly cameraSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly instances: readonly OverlayInstanceV3[]
}): OverlayBatchV3 {
  const packed = new Float32Array(input.instances.length * OVERLAY_INSTANCE_FLOATS_V3)
  let kindMask = 0
  input.instances.forEach((instance, index) => {
    const tag = KIND_TAGS[instance.kind]
    const color = parseColor(instance.color)
    kindMask |= 1 << tag
    const offset = index * OVERLAY_INSTANCE_FLOATS_V3
    packed[offset] = tag
    packed[offset + 1] = instance.x
    packed[offset + 2] = instance.y
    packed[offset + 3] = instance.width
    packed[offset + 4] = instance.height
    packed[offset + 5] = color.r
    packed[offset + 6] = color.g
    packed[offset + 7] = color.b
    packed[offset + 8] = Math.max(0, Math.min(1, instance.alpha ?? 1))
    packed[offset + 9] = instance.z ?? 0
  })
  return {
    axisSeqX: input.axisSeqX,
    axisSeqY: input.axisSeqY,
    cameraSeq: input.cameraSeq,
    count: input.instances.length,
    instances: packed,
    kindMask,
    seq: input.seq,
  }
}

function parseColor(color: string): { r: number; g: number; b: number } {
  if (!/^#[0-9a-fA-F]{6}$/.test(color)) {
    return { b: 0, g: 0, r: 0 }
  }
  return {
    b: Number.parseInt(color.slice(5, 7), 16) / 255,
    g: Number.parseInt(color.slice(3, 5), 16) / 255,
    r: Number.parseInt(color.slice(1, 3), 16) / 255,
  }
}
