import { isStaleValidGridTileKeyV2, serializeGridTileKeyV2, type GridScenePacketV2, type GridTileKeyV2 } from './scene-packet-v2.js'
import { validateGridScenePacketV2 } from './scene-packet-validator.js'
import { noteTypeGpuTileCacheEviction, noteTypeGpuTileCacheStaleLookup, noteTypeGpuTileCacheVisibleMark } from './grid-render-counters.js'

export interface TileGpuCacheEntry {
  readonly key: string
  readonly packet: GridScenePacketV2
  readonly visible: boolean
  readonly lastUsedSeq: number
}

interface MutableTileGpuCacheEntry extends TileGpuCacheEntry {
  packet: GridScenePacketV2
  visible: boolean
  lastUsedSeq: number
  compatibilityKey: string
}

export class TileGpuCache {
  private readonly entries = new Map<string, MutableTileGpuCacheEntry>()
  private readonly compatibilityBuckets = new Map<string, Set<MutableTileGpuCacheEntry>>()
  private readonly visibleKeys = new Set<string>()
  private seq = 0

  get size(): number {
    return this.entries.size
  }

  upsert(packet: GridScenePacketV2): TileGpuCacheEntry {
    const validation = validateGridScenePacketV2(packet)
    if (!validation.ok) {
      throw new Error(`Invalid grid scene packet: ${validation.reason}`)
    }
    const key = buildTileGpuCacheKey(packet)
    const existing = this.entries.get(key)
    if (existing && existing.packet.generation > packet.generation) {
      return existing
    }
    if (existing) {
      const compatibilityKey = buildTileGpuCompatibilityKey(packet.key)
      if (existing.compatibilityKey !== compatibilityKey) {
        this.removeFromCompatibilityBucket(existing)
        existing.compatibilityKey = compatibilityKey
        this.addToCompatibilityBucket(existing)
      }
      existing.packet = packet
      this.touch(existing)
      return existing
    }
    const entry: MutableTileGpuCacheEntry = {
      compatibilityKey: buildTileGpuCompatibilityKey(packet.key),
      key,
      lastUsedSeq: ++this.seq,
      packet,
      visible: false,
    }
    this.entries.set(key, entry)
    this.addToCompatibilityBucket(entry)
    return entry
  }

  findStaleValid(desiredKey: GridTileKeyV2, options?: { readonly excludeKey?: string | undefined }): TileGpuCacheEntry | null {
    const bucket = this.compatibilityBuckets.get(buildTileGpuCompatibilityKey(desiredKey))
    let match: TileGpuCacheEntry | null = null
    let scannedEntries = 0
    for (const entry of bucket ?? []) {
      scannedEntries += 1
      if (entry.key === options?.excludeKey || !isStaleValidGridTileKeyV2(entry.packet.key, desiredKey)) {
        continue
      }
      if (!match || entry.lastUsedSeq > match.lastUsedSeq) {
        match = entry
      }
    }
    noteTypeGpuTileCacheStaleLookup(scannedEntries, match !== null)
    if (!match) {
      return null
    }
    const mutableMatch = this.entries.get(match.key)
    if (!mutableMatch) {
      return match
    }
    this.touch(mutableMatch)
    return mutableMatch
  }

  get(key: string): TileGpuCacheEntry | null {
    const entry = this.entries.get(key) ?? null
    if (!entry) {
      return null
    }
    this.touch(entry)
    return entry
  }

  markVisible(keys: ReadonlySet<string>): void {
    for (const key of this.visibleKeys) {
      if (keys.has(key)) {
        continue
      }
      const entry = this.entries.get(key)
      if (entry) {
        entry.visible = false
      }
      this.visibleKeys.delete(key)
    }

    let marked = 0
    for (const key of keys) {
      const entry = this.entries.get(key)
      if (!entry) {
        continue
      }
      marked += 1
      entry.visible = true
      this.touch(entry)
      this.visibleKeys.add(key)
    }
    noteTypeGpuTileCacheVisibleMark(marked)
  }

  evictTo(maxEntries: number): void {
    const target = Math.max(0, Math.floor(maxEntries))
    if (this.entries.size <= target) {
      return
    }
    let evicted = 0
    while (this.entries.size > target) {
      const entry = this.findOldestEvictable()
      if (!entry) {
        if (evicted > 0) {
          noteTypeGpuTileCacheEviction(evicted)
        }
        return
      }
      this.deleteEntry(entry)
      evicted += 1
    }
    if (evicted > 0) {
      noteTypeGpuTileCacheEviction(evicted)
    }
  }

  private findOldestEvictable(): MutableTileGpuCacheEntry | null {
    for (const entry of this.entries.values()) {
      if (entry.visible) {
        continue
      }
      return entry
    }
    return null
  }

  private addToCompatibilityBucket(entry: MutableTileGpuCacheEntry): void {
    const bucket = this.compatibilityBuckets.get(entry.compatibilityKey) ?? new Set<MutableTileGpuCacheEntry>()
    bucket.add(entry)
    this.compatibilityBuckets.set(entry.compatibilityKey, bucket)
  }

  private removeFromCompatibilityBucket(entry: MutableTileGpuCacheEntry): void {
    const bucket = this.compatibilityBuckets.get(entry.compatibilityKey)
    if (!bucket) {
      return
    }
    bucket.delete(entry)
    if (bucket.size === 0) {
      this.compatibilityBuckets.delete(entry.compatibilityKey)
    }
  }

  private touch(entry: MutableTileGpuCacheEntry): void {
    entry.lastUsedSeq = ++this.seq
    this.entries.delete(entry.key)
    this.entries.set(entry.key, entry)
  }

  private deleteEntry(entry: MutableTileGpuCacheEntry): void {
    this.entries.delete(entry.key)
    this.visibleKeys.delete(entry.key)
    this.removeFromCompatibilityBucket(entry)
  }
}

export function syncTileGpuCacheFromPanes(input: {
  readonly cache: TileGpuCache
  readonly panes: readonly { readonly packedScene: GridScenePacketV2 }[]
  readonly maxEntries?: number | undefined
}): void {
  const visibleKeys = new Set<string>()
  for (const pane of input.panes) {
    visibleKeys.add(input.cache.upsert(pane.packedScene).key)
  }
  input.cache.markVisible(visibleKeys)
  input.cache.evictTo(input.maxEntries ?? 128)
}

export function buildTileGpuCacheKey(packet: GridScenePacketV2): string {
  return serializeGridTileKeyV2(packet.key)
}

function buildTileGpuCompatibilityKey(key: GridTileKeyV2): string {
  return [
    key.sheetName,
    key.paneKind,
    key.axisVersionX,
    key.axisVersionY,
    key.valueVersion,
    key.styleVersion,
    key.selectionIndependentVersion,
    key.freezeVersion,
    key.textEpoch,
    key.dprBucket,
  ].join(':')
}
