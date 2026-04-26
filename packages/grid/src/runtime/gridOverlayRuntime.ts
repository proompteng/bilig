import { packOverlayBatchV3, type OverlayBatchV3, type OverlayInstanceV3, type OverlayKindV3 } from '../renderer-v3/overlay-layer.js'

export class GridOverlayRuntime {
  private readonly instancesById = new Map<string, OverlayInstanceV3>()
  private seq = 0

  snapshot(): readonly OverlayInstanceV3[] {
    return [...this.instancesById.values()]
  }

  set(id: string, instance: OverlayInstanceV3): number {
    this.instancesById.set(id, instance)
    this.seq += 1
    return this.seq
  }

  delete(id: string): boolean {
    const deleted = this.instancesById.delete(id)
    if (deleted) {
      this.seq += 1
    }
    return deleted
  }

  clearKind(kind: OverlayKindV3): number {
    let changed = false
    for (const [id, instance] of this.instancesById) {
      if (instance.kind !== kind) {
        continue
      }
      this.instancesById.delete(id)
      changed = true
    }
    if (changed) {
      this.seq += 1
    }
    return this.seq
  }

  clear(): void {
    if (this.instancesById.size === 0) {
      return
    }
    this.instancesById.clear()
    this.seq += 1
  }

  buildBatch(input: { readonly cameraSeq: number; readonly axisSeqX: number; readonly axisSeqY: number }): OverlayBatchV3 {
    return packOverlayBatchV3({
      axisSeqX: input.axisSeqX,
      axisSeqY: input.axisSeqY,
      cameraSeq: input.cameraSeq,
      instances: this.snapshot(),
      seq: this.seq,
    })
  }
}
