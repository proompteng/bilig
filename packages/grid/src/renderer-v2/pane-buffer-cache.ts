import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { TextDecorationRect } from './line-text-quad-buffer.js'
import type { RectInstanceVertexBuffer, SurfaceUniformBuffer, TextInstanceVertexBuffer } from './typegpu-backend.js'
import type { TgpuBindGroup } from 'typegpu'

export interface WorkbookPaneBufferEntry {
  rectBuffer: RectInstanceVertexBuffer | null
  rectCapacity: number
  rectCount: number
  rectScene: GridGpuScene | null
  rectSignature: string | null
  surfaceUniform: SurfaceUniformBuffer | null
  surfaceBindGroup: TgpuBindGroup | null
  textBuffer: TextInstanceVertexBuffer | null
  textCapacity: number
  textCount: number
  textScene: GridTextScene | null
  textSignature: string | null
  textBindGroup: TgpuBindGroup | null
  textBindGroupAtlasVersion: number
  decorationRects: readonly TextDecorationRect[] | null
}

function createEmptyEntry(): WorkbookPaneBufferEntry {
  return {
    rectBuffer: null,
    rectCapacity: 0,
    rectCount: 0,
    rectScene: null,
    rectSignature: null,
    surfaceUniform: null,
    surfaceBindGroup: null,
    textBuffer: null,
    textCapacity: 0,
    textCount: 0,
    textScene: null,
    textSignature: null,
    textBindGroup: null,
    textBindGroupAtlasVersion: -1,
    decorationRects: null,
  }
}

export class WorkbookPaneBufferCache {
  private readonly entries = new Map<string, WorkbookPaneBufferEntry>()

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
    entry.rectBuffer?.destroy()
    entry.surfaceUniform?.buffer.destroy()
    entry.textBuffer?.destroy()
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
      this.delete(paneId)
    }
  }
}
