import { isStaleValidGridTileKeyV2, serializeGridTileKeyV2, type GridScenePacketV2, type GridTileKeyV2 } from './scene-packet-v2.js'
import { validateGridScenePacketV2 } from './scene-packet-validator.js'
import { noteTypeGpuTileCacheEviction, noteTypeGpuTileCacheStaleLookup, noteTypeGpuTileCacheVisibleMark } from './grid-render-counters.js'

export interface TileGpuCacheEntry {
  readonly key: string
  readonly packet: GridScenePacketV2
  readonly visible: boolean
  readonly lastUsedSeq: number
}

export class TileGpuCache {
  private readonly entries = new Map<string, TileGpuCacheEntry>()
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
    const entry = {
      key,
      packet,
      visible: existing?.visible ?? false,
      lastUsedSeq: ++this.seq,
    }
    this.entries.set(key, entry)
    return entry
  }

  findStaleValid(desiredKey: GridTileKeyV2, options?: { readonly excludeKey?: string | undefined }): TileGpuCacheEntry | null {
    let match: TileGpuCacheEntry | null = null
    let scannedEntries = 0
    for (const entry of this.entries.values()) {
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
    const next = { ...match, lastUsedSeq: ++this.seq }
    this.entries.set(match.key, next)
    return next
  }

  get(key: string): TileGpuCacheEntry | null {
    const entry = this.entries.get(key) ?? null
    if (!entry) {
      return null
    }
    const next = { ...entry, lastUsedSeq: ++this.seq }
    this.entries.set(key, next)
    return next
  }

  markVisible(keys: ReadonlySet<string>): void {
    let marked = 0
    for (const [key, entry] of this.entries) {
      const visible = keys.has(key)
      if (visible) {
        marked += 1
      }
      if (entry.visible === visible && !visible) {
        continue
      }
      this.entries.set(key, { ...entry, visible, lastUsedSeq: visible ? ++this.seq : entry.lastUsedSeq })
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
      this.entries.delete(entry.key)
      evicted += 1
    }
    if (evicted > 0) {
      noteTypeGpuTileCacheEviction(evicted)
    }
  }

  private findOldestEvictable(): TileGpuCacheEntry | null {
    let oldest: TileGpuCacheEntry | null = null
    for (const entry of this.entries.values()) {
      if (entry.visible) {
        continue
      }
      if (!oldest || entry.lastUsedSeq < oldest.lastUsedSeq) {
        oldest = entry
      }
    }
    return oldest
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
