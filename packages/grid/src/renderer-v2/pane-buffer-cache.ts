import type { TextDecorationRect } from './line-text-quad-buffer.js'
import type { RectInstanceVertexBuffer, SurfaceUniformBuffer, TextInstanceVertexBuffer } from './typegpu-backend.js'
import type { TgpuBindGroup } from 'typegpu'

export interface WorkbookPaneBufferEntry {
  rectBuffer: RectInstanceVertexBuffer | null
  rectCapacity: number
  rectCount: number
  rectSignature: string | null
  surfaceUniform: SurfaceUniformBuffer | null
  surfaceBindGroup: TgpuBindGroup | null
  textBuffer: TextInstanceVertexBuffer | null
  textCapacity: number
  textCount: number
  textSignature: string | null
  textBindGroup: TgpuBindGroup | null
  textBindGroupAtlasVersion: number
  decorationRects: readonly TextDecorationRect[] | null
}

interface ReleasedRectBuffer {
  readonly buffer: RectInstanceVertexBuffer
  readonly capacity: number
}

interface ReleasedTextBuffer {
  readonly buffer: TextInstanceVertexBuffer
  readonly capacity: number
}

function createEmptyEntry(): WorkbookPaneBufferEntry {
  return {
    rectBuffer: null,
    rectCapacity: 0,
    rectCount: 0,
    rectSignature: null,
    surfaceUniform: null,
    surfaceBindGroup: null,
    textBuffer: null,
    textCapacity: 0,
    textCount: 0,
    textSignature: null,
    textBindGroup: null,
    textBindGroupAtlasVersion: -1,
    decorationRects: null,
  }
}

export class WorkbookPaneBufferCache {
  private readonly entries = new Map<string, WorkbookPaneBufferEntry>()
  private readonly freeRectBuffers: ReleasedRectBuffer[] = []
  private readonly freeTextBuffers: ReleasedTextBuffer[] = []

  get(paneId: string): WorkbookPaneBufferEntry {
    const existing = this.entries.get(paneId)
    if (existing) {
      return existing
    }
    const next = createEmptyEntry()
    this.entries.set(paneId, next)
    return next
  }

  peek(paneId: string): WorkbookPaneBufferEntry | null {
    return this.entries.get(paneId) ?? null
  }

  delete(paneId: string): void {
    const entry = this.entries.get(paneId)
    if (!entry) {
      return
    }
    this.dropEntry(entry, { reuseVertexBuffers: true })
    this.entries.delete(paneId)
  }

  pruneExcept(paneIds: ReadonlySet<string>): void {
    for (const paneId of this.entries.keys()) {
      if (!paneIds.has(paneId)) {
        this.delete(paneId)
      }
    }
  }

  dispose(): void {
    for (const paneId of this.entries.keys()) {
      const entry = this.entries.get(paneId)
      if (entry) {
        this.dropEntry(entry, { reuseVertexBuffers: false })
      }
    }
    this.entries.clear()
    this.freeRectBuffers.forEach(({ buffer }) => buffer.destroy())
    this.freeTextBuffers.forEach(({ buffer }) => buffer.destroy())
    this.freeRectBuffers.length = 0
    this.freeTextBuffers.length = 0
  }

  acquireRectBuffer(minCapacity: number): ReleasedRectBuffer | null {
    return takeSmallestCapacity(this.freeRectBuffers, minCapacity)
  }

  releaseRectBuffer(buffer: RectInstanceVertexBuffer, capacity: number): void {
    this.freeRectBuffers.push({ buffer, capacity: Math.max(0, capacity) })
  }

  acquireTextBuffer(minCapacity: number): ReleasedTextBuffer | null {
    return takeSmallestCapacity(this.freeTextBuffers, minCapacity)
  }

  releaseTextBuffer(buffer: TextInstanceVertexBuffer, capacity: number): void {
    this.freeTextBuffers.push({ buffer, capacity: Math.max(0, capacity) })
  }

  private dropEntry(entry: WorkbookPaneBufferEntry, options: { readonly reuseVertexBuffers: boolean }): void {
    if (entry.rectBuffer) {
      if (options.reuseVertexBuffers) {
        this.releaseRectBuffer(entry.rectBuffer, entry.rectCapacity)
      } else {
        entry.rectBuffer.destroy()
      }
      entry.rectBuffer = null
      entry.rectCapacity = 0
      entry.rectCount = 0
    }
    entry.surfaceUniform?.buffer.destroy()
    entry.surfaceUniform = null
    entry.surfaceBindGroup = null
    if (entry.textBuffer) {
      if (options.reuseVertexBuffers) {
        this.releaseTextBuffer(entry.textBuffer, entry.textCapacity)
      } else {
        entry.textBuffer.destroy()
      }
      entry.textBuffer = null
      entry.textCapacity = 0
      entry.textCount = 0
    }
    entry.textBindGroup = null
    entry.textBindGroupAtlasVersion = -1
  }
}

function takeSmallestCapacity<TBuffer extends { readonly capacity: number }>(buffers: TBuffer[], minCapacity: number): TBuffer | null {
  let bestIndex = -1
  let bestCapacity = Number.POSITIVE_INFINITY
  const required = Math.max(1, Math.ceil(minCapacity))
  for (let index = 0; index < buffers.length; index += 1) {
    const candidate = buffers[index]
    if (!candidate || candidate.capacity < required || candidate.capacity >= bestCapacity) {
      continue
    }
    bestCapacity = candidate.capacity
    bestIndex = index
  }
  if (bestIndex < 0) {
    return null
  }
  const [buffer] = buffers.splice(bestIndex, 1)
  return buffer ?? null
}
