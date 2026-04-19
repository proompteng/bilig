import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'
import type { TextDecorationRect } from './text-quad-buffer.js'

export interface WorkbookPaneBufferEntry {
  rectBuffer: GPUBuffer | null
  rectCapacity: number
  rectCount: number
  rectScene: GridGpuScene | null
  textBuffer: GPUBuffer | null
  textCapacity: number
  textCount: number
  textScene: GridTextScene | null
  decorationRects: readonly TextDecorationRect[] | null
}

function createEmptyEntry(): WorkbookPaneBufferEntry {
  return {
    rectBuffer: null,
    rectCapacity: 0,
    rectCount: 0,
    rectScene: null,
    textBuffer: null,
    textCapacity: 0,
    textCount: 0,
    textScene: null,
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

  delete(paneId: string): void {
    const entry = this.entries.get(paneId)
    if (!entry) {
      return
    }
    entry.rectBuffer?.destroy()
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
