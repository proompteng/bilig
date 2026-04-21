import type { GridScenePacketV2 } from './scene-packet-v2.js'
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

export function buildTileGpuCacheKey(packet: GridScenePacketV2): string {
  const viewport = packet.viewport
  return [packet.sheetName, packet.paneId, viewport.rowStart, viewport.rowEnd, viewport.colStart, viewport.colEnd].join(':')
}
