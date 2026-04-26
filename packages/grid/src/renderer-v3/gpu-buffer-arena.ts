export type GpuBufferLayoutV3 = 'rectInstances' | 'textRuns' | 'overlayInstances' | 'axis' | 'uniform'

export interface GpuBufferHandleV3<Buffer = unknown> {
  readonly buffer: Buffer
  readonly layout: GpuBufferLayoutV3
  readonly capacityBytes: number
  usedBytes: number
  readonly classId: number
}

export interface GpuBufferArenaStatsV3 {
  readonly creates: number
  readonly reuses: number
  readonly destroys: number
  readonly freeBytes: number
}

export class GpuBufferArenaV3<Buffer = unknown> {
  private readonly freeLists = new Map<string, GpuBufferHandleV3<Buffer>[]>()
  private createdCount = 0
  private reusedCount = 0
  private destroyedCount = 0
  private freeByteCount = 0

  constructor(
    private readonly createBuffer: (input: {
      readonly layout: GpuBufferLayoutV3
      readonly capacityBytes: number
      readonly classId: number
    }) => Buffer,
    private readonly destroyBuffer: (buffer: Buffer) => void = () => undefined,
  ) {}

  stats(): GpuBufferArenaStatsV3 {
    return {
      creates: this.createdCount,
      destroys: this.destroyedCount,
      freeBytes: this.freeByteCount,
      reuses: this.reusedCount,
    }
  }

  acquire(layout: GpuBufferLayoutV3, requiredBytes: number): GpuBufferHandleV3<Buffer> {
    const capacityBytes = resolveCapacityBytes(requiredBytes)
    const classId = Math.log2(capacityBytes)
    const key = buildFreeListKey(layout, classId)
    const freeList = this.freeLists.get(key)
    const reused = freeList?.pop()
    if (reused) {
      this.freeByteCount -= reused.capacityBytes
      reused.usedBytes = Math.max(0, requiredBytes)
      this.reusedCount += 1
      return reused
    }
    this.createdCount += 1
    return {
      buffer: this.createBuffer({ capacityBytes, classId, layout }),
      capacityBytes,
      classId,
      layout,
      usedBytes: Math.max(0, requiredBytes),
    }
  }

  release(handle: GpuBufferHandleV3<Buffer>): void {
    handle.usedBytes = 0
    const key = buildFreeListKey(handle.layout, handle.classId)
    const freeList = this.freeLists.get(key) ?? []
    freeList.push(handle)
    this.freeLists.set(key, freeList)
    this.freeByteCount += handle.capacityBytes
  }

  trim(bytesToFree: number): number {
    let remaining = Math.max(0, bytesToFree)
    let freed = 0
    for (const [key, freeList] of this.freeLists) {
      while (remaining > 0 && freeList.length > 0) {
        const handle = freeList.pop()
        if (!handle) {
          break
        }
        this.destroyBuffer(handle.buffer)
        this.destroyedCount += 1
        this.freeByteCount -= handle.capacityBytes
        remaining -= handle.capacityBytes
        freed += handle.capacityBytes
      }
      if (freeList.length === 0) {
        this.freeLists.delete(key)
      }
      if (remaining <= 0) {
        break
      }
    }
    return freed
  }
}

function resolveCapacityBytes(requiredBytes: number): number {
  const required = Math.max(1, Math.ceil(requiredBytes))
  return 2 ** Math.max(8, Math.ceil(Math.log2(required)))
}

function buildFreeListKey(layout: GpuBufferLayoutV3, classId: number): string {
  return `${layout}:${classId}`
}
