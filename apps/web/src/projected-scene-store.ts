import type { ViewportPatch, WorkerEngineClient } from '@bilig/worker-transport'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
} from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import { residentPaneSceneRequestNeedsRefresh } from './projected-scene-damage.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'

interface SceneEntry {
  readonly key: string
  readonly request: WorkbookPaneSceneRequest
  listeners: Set<() => void>
  scenes: readonly WorkbookPaneScenePacket[] | null
  inFlight: Promise<void> | null
  refreshQueued: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isOptionalPackedScene(value: unknown): boolean {
  if (value === undefined) {
    return true
  }
  return isGridScenePacketV2(value) && validateGridScenePacketV2(value).ok
}

function isViewportRecord(value: unknown): value is GridScenePacketV2['viewport'] {
  return (
    isRecord(value) &&
    typeof value['rowStart'] === 'number' &&
    typeof value['rowEnd'] === 'number' &&
    typeof value['colStart'] === 'number' &&
    typeof value['colEnd'] === 'number'
  )
}

function isSurfaceSizeRecord(value: unknown): value is GridScenePacketV2['surfaceSize'] {
  return isRecord(value) && typeof value['width'] === 'number' && typeof value['height'] === 'number'
}

function isGridScenePacketV2(value: unknown): value is GridScenePacketV2 {
  return (
    isRecord(value) &&
    value['magic'] === GRID_SCENE_PACKET_V2_MAGIC &&
    value['version'] === GRID_SCENE_PACKET_V2_VERSION &&
    typeof value['generation'] === 'number' &&
    typeof value['sheetName'] === 'string' &&
    typeof value['paneId'] === 'string' &&
    isViewportRecord(value['viewport']) &&
    isSurfaceSizeRecord(value['surfaceSize']) &&
    value['rects'] instanceof Float32Array &&
    typeof value['rectCount'] === 'number' &&
    value['textMetrics'] instanceof Float32Array &&
    typeof value['textCount'] === 'number'
  )
}

function isDisposedWorkerClientError(error: unknown): boolean {
  return error instanceof Error && error.message === 'Worker engine client disposed'
}

function isResidentPaneScenePacketArray(value: unknown): value is readonly WorkbookPaneScenePacket[] {
  return (
    Array.isArray(value) &&
    value.every(
      (entry) =>
        isRecord(entry) &&
        typeof entry['generation'] === 'number' &&
        typeof entry['paneId'] === 'string' &&
        isRecord(entry['viewport']) &&
        isRecord(entry['surfaceSize']) &&
        isRecord(entry['gpuScene']) &&
        isRecord(entry['textScene']) &&
        isOptionalPackedScene(entry['packedScene']),
    )
  )
}

function buildRequestKey(request: WorkbookPaneSceneRequest): string {
  const range = request.selectionRange
  return [
    request.sheetName,
    request.residentViewport.rowStart,
    request.residentViewport.rowEnd,
    request.residentViewport.colStart,
    request.residentViewport.colEnd,
    request.freezeRows,
    request.freezeCols,
    request.selectedCell.col,
    request.selectedCell.row,
    range?.x ?? -1,
    range?.y ?? -1,
    range?.width ?? -1,
    range?.height ?? -1,
    request.editingCell?.col ?? -1,
    request.editingCell?.row ?? -1,
  ].join(':')
}

export class ProjectedSceneStore {
  private readonly entries = new Map<string, SceneEntry>()

  constructor(private readonly client?: Pick<WorkerEngineClient, 'invoke'>) {}

  peekResidentPaneScenes(request: WorkbookPaneSceneRequest): readonly WorkbookPaneScenePacket[] | null {
    return this.entries.get(buildRequestKey(request))?.scenes ?? null
  }

  subscribeResidentPaneScenes(request: WorkbookPaneSceneRequest, listener: () => void): () => void {
    if (!this.client) {
      throw new Error('Worker resident pane scene subscriptions require a worker engine client')
    }
    const key = buildRequestKey(request)
    const entry = this.entries.get(key) ?? {
      key,
      request,
      listeners: new Set<() => void>(),
      scenes: null,
      inFlight: null,
      refreshQueued: false,
    }
    entry.listeners.add(listener)
    this.entries.set(key, entry)
    if (entry.scenes === null) {
      this.queueRefresh(entry)
    }
    return () => {
      entry.listeners.delete(listener)
    }
  }

  dropSheets(sheetNames: readonly string[]): void {
    if (sheetNames.length === 0) {
      return
    }
    const removed = new Set(sheetNames)
    for (const [key, entry] of this.entries) {
      if (removed.has(entry.request.sheetName)) {
        this.entries.delete(key)
      }
    }
  }

  noteViewportPatch(patch: ViewportPatch): void {
    for (const entry of this.entries.values()) {
      if (!residentPaneSceneRequestNeedsRefresh(entry.request, patch)) {
        continue
      }
      this.queueRefresh(entry)
    }
  }

  private queueRefresh(entry: SceneEntry): void {
    if (entry.refreshQueued) {
      return
    }
    entry.refreshQueued = true
    if (entry.inFlight) {
      return
    }
    this.scheduleRefresh(entry)
  }

  private scheduleRefresh(entry: SceneEntry): void {
    const flush = () => {
      if (!entry.refreshQueued || entry.inFlight) {
        return
      }
      entry.refreshQueued = false
      void this.refreshEntry(entry)
    }
    if (typeof window === 'undefined') {
      setTimeout(flush, 0)
      return
    }
    window.requestAnimationFrame(flush)
  }

  private async refreshEntry(entry: SceneEntry): Promise<void> {
    const client = this.client
    if (!client) {
      return
    }
    if (entry.inFlight) {
      return
    }
    entry.inFlight = (async (): Promise<void> => {
      try {
        const isIncrementalRefresh = entry.scenes !== null
        const next = await client.invoke('getResidentPaneScenes', entry.request)
        if (!isResidentPaneScenePacketArray(next)) {
          throw new Error('Worker returned an unexpected resident pane scene payload')
        }
        entry.scenes = next
        if (isIncrementalRefresh) {
          getWorkbookScrollPerfCollector()?.noteScenePacketRefresh(next.length)
        }
        entry.listeners.forEach((listener) => listener())
      } catch (error) {
        if (entry.listeners.size === 0 || isDisposedWorkerClientError(error)) {
          return
        }
      }
    })()
    try {
      await entry.inFlight
    } finally {
      entry.inFlight = null
      if (entry.refreshQueued && entry.listeners.size > 0) {
        this.scheduleRefresh(entry)
      }
    }
  }
}
