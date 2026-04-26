import type { GlyphIdV3 } from './glyph-key.js'

export type TextRunWrapModeV3 = 'clip' | 'overflow' | 'wrap'

export interface TextRunKeyV3 {
  readonly textInternId: number
  readonly fontInternId: number
  readonly colorId: number
  readonly horizontalAlign: number
  readonly verticalAlign: number
  readonly wrapMode: TextRunWrapModeV3
  readonly clipWidthBucket: number
  readonly dprBucket: number
}

export interface TextRunRecordV3<Payload extends ArrayBufferView = Uint32Array | Float32Array> {
  readonly runId: number
  readonly key: TextRunKeyV3
  readonly glyphIds: readonly GlyphIdV3[]
  payload: Payload
  byteSize: number
  lastUsedGeneration: number
}

export interface TextRunCacheStatsV3 {
  readonly runCount: number
  readonly byteSize: number
  readonly generation: number
}

export class TextRunCacheV3<Payload extends ArrayBufferView = Uint32Array | Float32Array> {
  private readonly recordsByKey = new Map<string, TextRunRecordV3<Payload>>()
  private readonly recordsById = new Map<number, TextRunRecordV3<Payload>>()
  private nextRunId = 1
  private generation = 0
  private bytes = 0

  stats(): TextRunCacheStatsV3 {
    return {
      byteSize: this.bytes,
      generation: this.generation,
      runCount: this.recordsByKey.size,
    }
  }

  get(key: TextRunKeyV3): TextRunRecordV3<Payload> | null {
    const record = this.recordsByKey.get(encodeTextRunKeyV3(key)) ?? null
    if (!record) {
      return null
    }
    this.touch(record)
    return record
  }

  getById(runId: number): TextRunRecordV3<Payload> | null {
    const record = this.recordsById.get(runId) ?? null
    if (!record) {
      return null
    }
    this.touch(record)
    return record
  }

  put(input: { readonly key: TextRunKeyV3; readonly glyphIds: readonly GlyphIdV3[]; readonly payload: Payload }): TextRunRecordV3<Payload> {
    const encoded = encodeTextRunKeyV3(input.key)
    const existing = this.recordsByKey.get(encoded)
    const byteSize = input.payload.byteLength
    if (existing) {
      this.bytes += byteSize - existing.byteSize
      existing.payload = input.payload
      existing.byteSize = byteSize
      this.touch(existing)
      return existing
    }

    const record: TextRunRecordV3<Payload> = {
      byteSize,
      glyphIds: [...input.glyphIds],
      key: { ...input.key },
      lastUsedGeneration: 0,
      payload: input.payload,
      runId: this.nextRunId++,
    }
    this.recordsByKey.set(encoded, record)
    this.recordsById.set(record.runId, record)
    this.bytes += byteSize
    this.touch(record)
    return record
  }

  getOrCreate(input: {
    readonly key: TextRunKeyV3
    readonly create: () => { readonly glyphIds: readonly GlyphIdV3[]; readonly payload: Payload }
  }): TextRunRecordV3<Payload> {
    const existing = this.get(input.key)
    if (existing) {
      return existing
    }
    const created = input.create()
    return this.put({ glyphIds: created.glyphIds, key: input.key, payload: created.payload })
  }

  delete(runId: number): TextRunRecordV3<Payload> | null {
    const record = this.recordsById.get(runId) ?? null
    if (!record) {
      return null
    }
    this.recordsById.delete(record.runId)
    this.recordsByKey.delete(encodeTextRunKeyV3(record.key))
    this.bytes -= record.byteSize
    return record
  }

  evictToBudget(maxBytes: number, onEvict: (record: TextRunRecordV3<Payload>) => void = () => undefined): number {
    const budget = Math.max(0, maxBytes)
    let evicted = 0
    while (this.bytes > budget && this.recordsById.size > 0) {
      const victim = this.findLeastRecentlyUsed()
      if (!victim) {
        break
      }
      const removed = this.delete(victim.runId)
      if (!removed) {
        break
      }
      onEvict(removed)
      evicted += 1
    }
    return evicted
  }

  private touch(record: TextRunRecordV3<Payload>): void {
    this.generation += 1
    record.lastUsedGeneration = this.generation
  }

  private findLeastRecentlyUsed(): TextRunRecordV3<Payload> | null {
    let victim: TextRunRecordV3<Payload> | null = null
    for (const record of this.recordsById.values()) {
      if (!victim || record.lastUsedGeneration < victim.lastUsedGeneration) {
        victim = record
      }
    }
    return victim
  }
}

export function encodeTextRunKeyV3(key: TextRunKeyV3): string {
  return [
    key.textInternId,
    key.fontInternId,
    key.colorId,
    key.horizontalAlign,
    key.verticalAlign,
    key.wrapMode,
    key.clipWidthBucket,
    key.dprBucket,
  ].join(':')
}
