import type { TileKey53, TileKeyFields } from './tile-key.js'

export type TileEntryStateV3 = 'empty' | 'materializing' | 'ready' | 'dirty' | 'evicting'

export interface TileRevisionTupleV3 {
  readonly valueSeq: number
  readonly styleSeq: number
  readonly textSeq: number
  readonly rectSeq: number
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
}

export interface TileEntryV3<Packet = unknown, Resources = unknown> extends TileKeyFields, TileRevisionTupleV3 {
  readonly key: TileKey53
  compatibilityKey: number
  packet: Packet | null
  resources: Resources | null
  dirtyMask: number
  state: TileEntryStateV3
  byteSizeCpu: number
  byteSizeGpu: number
  lastUsedGeneration: number
  visibleGeneration: number
  pinnedUntilGeneration: number
  lruPrev: TileEntryV3<Packet, Resources> | null
  lruNext: TileEntryV3<Packet, Resources> | null
}

export interface TileUpsertInputV3<Packet = unknown, Resources = unknown> extends TileKeyFields, TileRevisionTupleV3 {
  readonly key: TileKey53
  readonly packet?: Packet | null | undefined
  readonly resources?: Resources | null | undefined
  readonly dirtyMask?: number | undefined
  readonly state?: TileEntryStateV3 | undefined
  readonly byteSizeCpu?: number | undefined
  readonly byteSizeGpu?: number | undefined
}

interface CompatibilityNode {
  readonly key: number
  readonly children: Map<number, CompatibilityNode>
}

export class TileResidencyV3<Packet = unknown, Resources = unknown> {
  private readonly exact = new Map<TileKey53, TileEntryV3<Packet, Resources>>()
  private readonly compatibilityBuckets = new Map<number, Set<TileEntryV3<Packet, Resources>>>()
  private readonly compatibilityRoot: CompatibilityNode = { key: 0, children: new Map() }
  private nextCompatibilityKey = 1
  private lruHead: TileEntryV3<Packet, Resources> | null = null
  private lruTail: TileEntryV3<Packet, Resources> | null = null
  private generation = 0
  private cpuBytes = 0
  private gpuBytes = 0
  private lastStaleScanCount = 0

  get size(): number {
    return this.exact.size
  }

  get bytesCpu(): number {
    return this.cpuBytes
  }

  get bytesGpu(): number {
    return this.gpuBytes
  }

  get currentGeneration(): number {
    return this.generation
  }

  getLastStaleScanCount(): number {
    return this.lastStaleScanCount
  }

  entries(): IterableIterator<TileEntryV3<Packet, Resources>> {
    return this.exact.values()
  }

  upsert(input: TileUpsertInputV3<Packet, Resources>): TileEntryV3<Packet, Resources> {
    const compatibilityKey = this.resolveCompatibilityKey(input)
    const existing = this.exact.get(input.key)
    if (existing) {
      this.cpuBytes += (input.byteSizeCpu ?? existing.byteSizeCpu) - existing.byteSizeCpu
      this.gpuBytes += (input.byteSizeGpu ?? existing.byteSizeGpu) - existing.byteSizeGpu
      if (existing.compatibilityKey !== compatibilityKey) {
        this.removeFromCompatibilityBucket(existing)
        existing.compatibilityKey = compatibilityKey
        this.addToCompatibilityBucket(existing)
      }
      existing.packet = input.packet ?? existing.packet
      existing.resources = input.resources ?? existing.resources
      existing.dirtyMask = input.dirtyMask ?? existing.dirtyMask
      existing.state = input.state ?? existing.state
      existing.byteSizeCpu = input.byteSizeCpu ?? existing.byteSizeCpu
      existing.byteSizeGpu = input.byteSizeGpu ?? existing.byteSizeGpu
      existing.lastUsedGeneration = ++this.generation
      this.moveToLruHead(existing)
      return existing
    }

    const entry: TileEntryV3<Packet, Resources> = {
      axisSeqX: input.axisSeqX,
      axisSeqY: input.axisSeqY,
      byteSizeCpu: input.byteSizeCpu ?? 0,
      byteSizeGpu: input.byteSizeGpu ?? 0,
      colTile: input.colTile,
      compatibilityKey,
      dirtyMask: input.dirtyMask ?? 0,
      dprBucket: input.dprBucket,
      freezeSeq: input.freezeSeq,
      key: input.key,
      lastUsedGeneration: ++this.generation,
      lruNext: null,
      lruPrev: null,
      packet: input.packet ?? null,
      pinnedUntilGeneration: 0,
      rectSeq: input.rectSeq,
      resources: input.resources ?? null,
      rowTile: input.rowTile,
      sheetOrdinal: input.sheetOrdinal,
      state: input.state ?? 'ready',
      styleSeq: input.styleSeq,
      textSeq: input.textSeq,
      valueSeq: input.valueSeq,
      visibleGeneration: 0,
    }
    this.exact.set(entry.key, entry)
    this.cpuBytes += entry.byteSizeCpu
    this.gpuBytes += entry.byteSizeGpu
    this.addToCompatibilityBucket(entry)
    this.insertLruHead(entry)
    return entry
  }

  getExact(key: TileKey53): TileEntryV3<Packet, Resources> | null {
    const entry = this.exact.get(key) ?? null
    if (!entry) {
      return null
    }
    entry.lastUsedGeneration = ++this.generation
    this.moveToLruHead(entry)
    return entry
  }

  findStaleCompatible(
    input: TileKeyFields &
      Pick<TileRevisionTupleV3, 'axisSeqX' | 'axisSeqY' | 'freezeSeq'> & {
        readonly excludeKey?: TileKey53 | undefined
      },
  ): TileEntryV3<Packet, Resources> | null {
    const compatibilityKey = this.resolveCompatibilityKey(input)
    const bucket = this.compatibilityBuckets.get(compatibilityKey)
    let scanned = 0
    let match: TileEntryV3<Packet, Resources> | null = null
    if (bucket) {
      for (const entry of bucket) {
        scanned += 1
        if (entry.key === input.excludeKey || entry.dirtyMask !== 0 || entry.state === 'evicting') {
          continue
        }
        if (!match || entry.lastUsedGeneration > match.lastUsedGeneration) {
          match = entry
        }
      }
    }
    this.lastStaleScanCount = scanned
    if (!match) {
      return null
    }
    match.lastUsedGeneration = ++this.generation
    this.moveToLruHead(match)
    return match
  }

  markVisible(keys: Iterable<TileKey53>): number {
    this.generation += 1
    let count = 0
    for (const key of keys) {
      const entry = this.exact.get(key)
      if (!entry) {
        continue
      }
      entry.visibleGeneration = this.generation
      entry.lastUsedGeneration = this.generation
      this.moveToLruHead(entry)
      count += 1
    }
    return count
  }

  isVisible(entry: TileEntryV3<Packet, Resources>): boolean {
    return entry.visibleGeneration === this.generation
  }

  pin(key: TileKey53, generations = 1): boolean {
    const entry = this.exact.get(key)
    if (!entry) {
      return false
    }
    entry.pinnedUntilGeneration = Math.max(entry.pinnedUntilGeneration, this.generation + Math.max(1, generations))
    return true
  }

  evictToBudgets(input: {
    readonly maxCpuBytes: number
    readonly maxGpuBytes: number
    readonly onEvict?: ((entry: TileEntryV3<Packet, Resources>) => void) | undefined
  }): number {
    const maxCpuBytes = Math.max(0, input.maxCpuBytes)
    const maxGpuBytes = Math.max(0, input.maxGpuBytes)
    let evicted = 0
    while ((this.cpuBytes > maxCpuBytes || this.gpuBytes > maxGpuBytes) && this.exact.size > 0) {
      const victim = this.findLruEvictionCandidate()
      if (!victim) {
        break
      }
      input.onEvict?.(victim)
      this.delete(victim.key)
      evicted += 1
    }
    return evicted
  }

  evictToSize(maxEntries: number, onEvict?: (entry: TileEntryV3<Packet, Resources>) => void): number {
    const targetSize = Math.max(0, maxEntries)
    let evicted = 0
    while (this.exact.size > targetSize) {
      const victim = this.findLruEvictionCandidate()
      if (!victim) {
        break
      }
      onEvict?.(victim)
      this.delete(victim.key)
      evicted += 1
    }
    return evicted
  }

  delete(key: TileKey53): boolean {
    const entry = this.exact.get(key)
    if (!entry) {
      return false
    }
    this.exact.delete(key)
    this.cpuBytes -= entry.byteSizeCpu
    this.gpuBytes -= entry.byteSizeGpu
    this.removeFromCompatibilityBucket(entry)
    this.removeFromLru(entry)
    return true
  }

  clear(): void {
    this.exact.clear()
    this.compatibilityBuckets.clear()
    this.compatibilityRoot.children.clear()
    this.nextCompatibilityKey = 1
    this.lruHead = null
    this.lruTail = null
    this.generation = 0
    this.cpuBytes = 0
    this.gpuBytes = 0
    this.lastStaleScanCount = 0
  }

  private findLruEvictionCandidate(): TileEntryV3<Packet, Resources> | null {
    let entry = this.lruTail
    while (entry) {
      if (!this.isVisible(entry) && entry.pinnedUntilGeneration < this.generation) {
        return entry
      }
      entry = entry.lruPrev
    }
    return null
  }

  private resolveCompatibilityKey(input: TileKeyFields & Pick<TileRevisionTupleV3, 'axisSeqX' | 'axisSeqY' | 'freezeSeq'>): number {
    let node = this.compatibilityRoot
    for (const value of [input.sheetOrdinal, input.dprBucket, input.axisSeqX, input.axisSeqY, input.freezeSeq]) {
      let child = node.children.get(value)
      if (!child) {
        child = { key: this.nextCompatibilityKey++, children: new Map() }
        node.children.set(value, child)
      }
      node = child
    }
    return node.key
  }

  private addToCompatibilityBucket(entry: TileEntryV3<Packet, Resources>): void {
    const bucket = this.compatibilityBuckets.get(entry.compatibilityKey) ?? new Set<TileEntryV3<Packet, Resources>>()
    bucket.add(entry)
    this.compatibilityBuckets.set(entry.compatibilityKey, bucket)
  }

  private removeFromCompatibilityBucket(entry: TileEntryV3<Packet, Resources>): void {
    const bucket = this.compatibilityBuckets.get(entry.compatibilityKey)
    if (!bucket) {
      return
    }
    bucket.delete(entry)
    if (bucket.size === 0) {
      this.compatibilityBuckets.delete(entry.compatibilityKey)
    }
  }

  private insertLruHead(entry: TileEntryV3<Packet, Resources>): void {
    entry.lruPrev = null
    entry.lruNext = this.lruHead
    if (this.lruHead) {
      this.lruHead.lruPrev = entry
    }
    this.lruHead = entry
    this.lruTail ??= entry
  }

  private moveToLruHead(entry: TileEntryV3<Packet, Resources>): void {
    if (this.lruHead === entry) {
      return
    }
    this.removeFromLru(entry)
    this.insertLruHead(entry)
  }

  private removeFromLru(entry: TileEntryV3<Packet, Resources>): void {
    if (entry.lruPrev) {
      entry.lruPrev.lruNext = entry.lruNext
    }
    if (entry.lruNext) {
      entry.lruNext.lruPrev = entry.lruPrev
    }
    if (this.lruHead === entry) {
      this.lruHead = entry.lruNext
    }
    if (this.lruTail === entry) {
      this.lruTail = entry.lruPrev
    }
    entry.lruPrev = null
    entry.lruNext = null
  }
}
