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

interface TextRunCacheRecordV3<Payload extends ArrayBufferView = Uint32Array | Float32Array> extends TextRunRecordV3<Payload> {
  glyphIds: GlyphIdV3[]
  lruPrev: TextRunCacheRecordV3<Payload> | null
  lruNext: TextRunCacheRecordV3<Payload> | null
}

export interface TextRunCacheStatsV3 {
  readonly runCount: number
  readonly byteSize: number
  readonly generation: number
}

export class TextRunCacheV3<Payload extends ArrayBufferView = Uint32Array | Float32Array> {
  private readonly recordsByKey = new Map<string, TextRunCacheRecordV3<Payload>>()
  private readonly recordsById = new Map<number, TextRunCacheRecordV3<Payload>>()
  private nextRunId = 1
  private generation = 0
  private bytes = 0
  private lruHead: TextRunCacheRecordV3<Payload> | null = null
  private lruTail: TextRunCacheRecordV3<Payload> | null = null

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
      existing.glyphIds = [...input.glyphIds]
      existing.payload = input.payload
      existing.byteSize = byteSize
      this.touch(existing)
      return existing
    }

    const record: TextRunCacheRecordV3<Payload> = {
      byteSize,
      glyphIds: [...input.glyphIds],
      key: { ...input.key },
      lastUsedGeneration: 0,
      lruNext: null,
      lruPrev: null,
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
    this.detach(record)
    return record
  }

  evictToBudget(maxBytes: number, onEvict: (record: TextRunRecordV3<Payload>) => void = () => undefined): number {
    const budget = Math.max(0, maxBytes)
    let evicted = 0
    while (this.bytes > budget && this.recordsById.size > 0) {
      const victim = this.lruHead
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

  private touch(record: TextRunCacheRecordV3<Payload>): void {
    this.generation += 1
    record.lastUsedGeneration = this.generation
    this.detach(record)
    if (this.lruTail) {
      this.lruTail.lruNext = record
      record.lruPrev = this.lruTail
      record.lruNext = null
      this.lruTail = record
      return
    }
    record.lruPrev = null
    record.lruNext = null
    this.lruHead = record
    this.lruTail = record
  }

  private detach(record: TextRunCacheRecordV3<Payload>): void {
    if (record.lruPrev) {
      record.lruPrev.lruNext = record.lruNext
    } else if (this.lruHead === record) {
      this.lruHead = record.lruNext
    }
    if (record.lruNext) {
      record.lruNext.lruPrev = record.lruPrev
    } else if (this.lruTail === record) {
      this.lruTail = record.lruPrev
    }
    record.lruPrev = null
    record.lruNext = null
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
