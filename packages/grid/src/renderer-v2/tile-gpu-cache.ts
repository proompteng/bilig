import { isStaleValidGridTileKeyV2, serializeGridTileKeyV2, type GridScenePacketV2, type GridTileKeyV2 } from './scene-packet-v2.js'
import { validateGridScenePacketV2 } from './scene-packet-validator.js'

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

  findStaleValid(desiredKey: GridTileKeyV2): TileGpuCacheEntry | null {
    const match = [...this.entries.values()]
      .filter((entry) => isStaleValidGridTileKeyV2(entry.packet.key, desiredKey))
      .toSorted((left, right) => right.lastUsedSeq - left.lastUsedSeq)[0]
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
    for (const [key, entry] of this.entries) {
      this.entries.set(key, { ...entry, visible: keys.has(key), lastUsedSeq: keys.has(key) ? ++this.seq : entry.lastUsedSeq })
    }
  }

  evictTo(maxEntries: number): void {
    const target = Math.max(0, Math.floor(maxEntries))
    if (this.entries.size <= target) {
      return
    }
    const evictable = [...this.entries.values()]
      .filter((entry) => !entry.visible)
      .toSorted((left, right) => left.lastUsedSeq - right.lastUsedSeq)
    for (const entry of evictable) {
      if (this.entries.size <= target) {
        return
      }
      this.entries.delete(entry.key)
    }
  }
}

export function syncTileGpuCacheFromPanes(input: {
  readonly cache: TileGpuCache
  readonly panes: readonly { readonly packedScene?: GridScenePacketV2 | undefined }[]
  readonly maxEntries?: number | undefined
}): void {
  const visibleKeys = new Set<string>()
  for (const pane of input.panes) {
    if (!pane.packedScene) {
      continue
    }
    visibleKeys.add(input.cache.upsert(pane.packedScene).key)
  }
  input.cache.markVisible(visibleKeys)
  input.cache.evictTo(input.maxEntries ?? 128)
}

export function buildTileGpuCacheKey(packet: GridScenePacketV2): string {
  return serializeGridTileKeyV2(packet.key)
}
