import type { ViewportPatch, WorkerEngineClient } from '@bilig/worker-transport'
import {
  GRID_SCENE_PACKET_V2_MAGIC,
  GRID_SCENE_PACKET_V2_VERSION,
  type GridScenePacketV2,
  type GridTileKeyV2,
} from '../../../packages/grid/src/renderer-v2/scene-packet-v2.js'
import { validateGridScenePacketV2 } from '../../../packages/grid/src/renderer-v2/scene-packet-validator.js'
import { residentPaneSceneRequestNeedsRefresh, type ResidentScenePatchDamage } from './projected-scene-damage.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import type { WorkbookPaneScenePacket, WorkbookPaneSceneRequest } from './resident-pane-scene-types.js'

interface SceneEntry {
  readonly key: string
  request: WorkbookPaneSceneRequest
  listeners: Set<() => void>
  scenes: readonly WorkbookPaneScenePacket[] | null
  coverageScenes: readonly WorkbookPaneScenePacket[] | null
  inFlight: Promise<void> | null
  refreshQueued: boolean
  refreshDelayMs: number
  refreshTimer: ReturnType<typeof setTimeout> | null
  activeRequestSeq: number
  sceneRevision: number
}

interface RetainedSceneEntry {
  readonly request: WorkbookPaneSceneRequest
  readonly scenes: readonly WorkbookPaneScenePacket[]
  readonly source: 'immediate' | 'worker'
  readonly lastUsedSeq: number
}

interface ProjectedSceneStoreOptions {
  readonly buildImmediateResidentPaneScenes?: (
    request: WorkbookPaneSceneRequest,
    generation: number,
  ) => readonly WorkbookPaneScenePacket[] | null
}

const BACKGROUND_WORKER_SCENE_REFRESH_DELAY_MS = 64
const MAX_RETAINED_RESIDENT_SCENE_REQUESTS = 32

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isPackedScene(value: unknown): boolean {
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

function isGridTileKeyV2(value: unknown): value is GridTileKeyV2 {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['paneKind'] === 'string' &&
    typeof value['rowStart'] === 'number' &&
    typeof value['rowEnd'] === 'number' &&
    typeof value['colStart'] === 'number' &&
    typeof value['colEnd'] === 'number' &&
    typeof value['rowTile'] === 'number' &&
    typeof value['colTile'] === 'number' &&
    typeof value['axisVersionX'] === 'number' &&
    typeof value['axisVersionY'] === 'number' &&
    typeof value['valueVersion'] === 'number' &&
    typeof value['styleVersion'] === 'number' &&
    typeof value['selectionIndependentVersion'] === 'number' &&
    typeof value['freezeVersion'] === 'number' &&
    typeof value['textEpoch'] === 'number' &&
    typeof value['dprBucket'] === 'number'
  )
}

function isGridScenePacketV2(value: unknown): value is GridScenePacketV2 {
  return (
    isRecord(value) &&
    value['magic'] === GRID_SCENE_PACKET_V2_MAGIC &&
    value['version'] === GRID_SCENE_PACKET_V2_VERSION &&
    typeof value['generation'] === 'number' &&
    typeof value['requestSeq'] === 'number' &&
    typeof value['cameraSeq'] === 'number' &&
    typeof value['generatedAt'] === 'number' &&
    isGridTileKeyV2(value['key']) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['paneId'] === 'string' &&
    isViewportRecord(value['viewport']) &&
    isSurfaceSizeRecord(value['surfaceSize']) &&
    value['rects'] instanceof Float32Array &&
    value['rectInstances'] instanceof Float32Array &&
    typeof value['rectCount'] === 'number' &&
    typeof value['fillRectCount'] === 'number' &&
    typeof value['borderRectCount'] === 'number' &&
    typeof value['rectSignature'] === 'string' &&
    value['textMetrics'] instanceof Float32Array &&
    Array.isArray(value['textRuns']) &&
    typeof value['textSignature'] === 'string' &&
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
        isPackedScene(entry['packedScene']),
    )
  )
}

function validateScenePacketResponseForRequest(
  scenes: readonly WorkbookPaneScenePacket[],
  request: WorkbookPaneSceneRequest,
): string | null {
  const expectedRequestSeq = request.requestSeq ?? 0
  const expectedCameraSeq = request.cameraSeq ?? expectedRequestSeq
  for (const scene of scenes) {
    const packet = scene.packedScene
    if (!packet) {
      return 'missing-packed-scene'
    }
    if (packet.requestSeq < expectedRequestSeq) {
      return 'request-sequence-stale'
    }
    if (packet.requestSeq !== expectedRequestSeq) {
      return 'request-sequence-mismatch'
    }
    if (packet.cameraSeq < expectedCameraSeq) {
      return 'camera-sequence-stale'
    }
    if (packet.sheetName !== request.sheetName) {
      return 'sheet-mismatch'
    }
  }
  return null
}

function buildRequestKey(request: WorkbookPaneSceneRequest): string {
  return [
    request.sheetName,
    request.residentViewport.rowStart,
    request.residentViewport.rowEnd,
    request.residentViewport.colStart,
    request.residentViewport.colEnd,
    request.freezeRows,
    request.freezeCols,
    request.dprBucket ?? 1,
  ].join(':')
}

function buildRetainedRequestKey(request: WorkbookPaneSceneRequest): string {
  return [
    request.sheetName,
    request.residentViewport.rowStart,
    request.residentViewport.rowEnd,
    request.residentViewport.colStart,
    request.residentViewport.colEnd,
    request.freezeRows,
    request.freezeCols,
    request.dprBucket ?? 1,
  ].join(':')
}

function resolveSceneRequestPriority(request: WorkbookPaneSceneRequest): number {
  if (Number.isInteger(request.priority) && request.priority !== undefined && request.priority >= 0) {
    return request.priority
  }
  return request.reason === 'prefetch' ? 4 : 0
}

function preferSceneRequest(current: WorkbookPaneSceneRequest, next: WorkbookPaneSceneRequest): WorkbookPaneSceneRequest {
  const currentPriority = resolveSceneRequestPriority(current)
  const nextPriority = resolveSceneRequestPriority(next)
  if (nextPriority < currentPriority) {
    return next
  }
  if (nextPriority === currentPriority && (next.cameraSeq ?? 0) > (current.cameraSeq ?? 0)) {
    return next
  }
  return current
}

function withRequestMetadata(request: WorkbookPaneSceneRequest, requestSeq: number, sceneRevision: number): WorkbookPaneSceneRequest {
  return {
    ...request,
    cameraSeq: request.cameraSeq ?? requestSeq,
    priority: resolveSceneRequestPriority(request),
    reason: request.reason ?? 'visible',
    requestSeq,
    sceneRevision,
  }
}

function resolveViewportPatchBatchId(patch: ViewportPatch): number | null {
  const batchId = patch.metrics?.batchId
  return Number.isInteger(batchId) && batchId !== undefined && batchId > 0 ? batchId : null
}

function viewportPatchHasStructuralDamage(patch: ViewportPatch, applied?: ResidentScenePatchDamage): boolean {
  if (applied?.axisChanged || applied?.freezeChanged) {
    return true
  }
  return patch.freezeRows !== undefined || patch.freezeCols !== undefined || patch.columns.length > 0 || patch.rows.length > 0
}

function residentScenesCoverViewportPatch(
  scenes: readonly WorkbookPaneScenePacket[] | null,
  patch: ViewportPatch,
  applied?: ResidentScenePatchDamage,
): boolean {
  if (!patch.full || viewportPatchHasStructuralDamage(patch, applied)) {
    return false
  }
  const batchId = resolveViewportPatchBatchId(patch)
  if (batchId === null || !scenes || scenes.length === 0) {
    return false
  }
  return scenes.every((scene) => {
    const key = scene.packedScene.key
    return key.valueVersion >= batchId && key.styleVersion >= batchId && key.selectionIndependentVersion >= batchId
  })
}

function sameViewport(
  left: Pick<GridScenePacketV2['viewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
  right: Pick<GridScenePacketV2['viewport'], 'rowStart' | 'rowEnd' | 'colStart' | 'colEnd'>,
): boolean {
  return (
    left.rowStart === right.rowStart && left.rowEnd === right.rowEnd && left.colStart === right.colStart && left.colEnd === right.colEnd
  )
}

function sameSurfaceSize(left: GridScenePacketV2['surfaceSize'], right: GridScenePacketV2['surfaceSize']): boolean {
  return left.width === right.width && left.height === right.height
}

function sameResidentScenePacketDrawSemantics(left: WorkbookPaneScenePacket, right: WorkbookPaneScenePacket): boolean {
  const leftPacket = left.packedScene
  const rightPacket = right.packedScene
  return (
    left.paneId === right.paneId &&
    leftPacket.sheetName === rightPacket.sheetName &&
    leftPacket.paneId === rightPacket.paneId &&
    leftPacket.key.sheetName === rightPacket.key.sheetName &&
    leftPacket.key.paneKind === rightPacket.key.paneKind &&
    leftPacket.key.rowStart === rightPacket.key.rowStart &&
    leftPacket.key.rowEnd === rightPacket.key.rowEnd &&
    leftPacket.key.colStart === rightPacket.key.colStart &&
    leftPacket.key.colEnd === rightPacket.key.colEnd &&
    leftPacket.key.rowTile === rightPacket.key.rowTile &&
    leftPacket.key.colTile === rightPacket.key.colTile &&
    leftPacket.key.dprBucket === rightPacket.key.dprBucket &&
    sameViewport(left.viewport, right.viewport) &&
    sameSurfaceSize(left.surfaceSize, right.surfaceSize) &&
    leftPacket.rectSignature === rightPacket.rectSignature &&
    leftPacket.textSignature === rightPacket.textSignature &&
    leftPacket.rectCount === rightPacket.rectCount &&
    leftPacket.fillRectCount === rightPacket.fillRectCount &&
    leftPacket.borderRectCount === rightPacket.borderRectCount &&
    leftPacket.textCount === rightPacket.textCount &&
    sameViewport(leftPacket.viewport, rightPacket.viewport) &&
    sameSurfaceSize(leftPacket.surfaceSize, rightPacket.surfaceSize)
  )
}

function sameResidentSceneDrawSemantics(
  left: readonly WorkbookPaneScenePacket[] | null,
  right: readonly WorkbookPaneScenePacket[],
): boolean {
  if (!left || left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    const leftScene = left[index]
    const rightScene = right[index]
    if (!leftScene || !rightScene || !sameResidentScenePacketDrawSemantics(leftScene, rightScene)) {
      return false
    }
  }
  return true
}

export class ProjectedSceneStore {
  private readonly entries = new Map<string, SceneEntry>()
  private readonly retainedScenes = new Map<string, RetainedSceneEntry>()
  private retainedSceneSeq = 0
  private nextRequestSeq = 0
  private nextImmediateSceneGeneration = 0
  private lastRejectedPacketReason: string | null = null

  constructor(
    private readonly client?: Pick<WorkerEngineClient, 'invoke'>,
    private readonly options: ProjectedSceneStoreOptions = {},
  ) {}

  getLastRejectedPacketReason(): string | null {
    return this.lastRejectedPacketReason
  }

  peekResidentPaneScenes(request: WorkbookPaneSceneRequest): readonly WorkbookPaneScenePacket[] | null {
    const entry = this.entries.get(buildRequestKey(request))
    if (entry) {
      return entry.scenes
    }
    return this.peekRetainedScenes(request)?.scenes ?? this.buildAndRetainImmediateScenes(request)
  }

  subscribeResidentPaneScenes(request: WorkbookPaneSceneRequest, listener: () => void): () => void {
    if (!this.client) {
      throw new Error('Worker resident pane scene subscriptions require a worker engine client')
    }
    const key = buildRequestKey(request)
    const existing = this.entries.get(key)
    let retained = existing ? null : this.peekRetainedScenes(request)
    if (!existing && !retained) {
      retained = this.buildAndRetainImmediateSceneEntry(request)
    }
    const entry =
      existing ??
      ({
        key,
        request,
        listeners: new Set<() => void>(),
        scenes: retained?.scenes ?? null,
        coverageScenes: retained?.scenes ?? null,
        inFlight: null,
        refreshQueued: false,
        refreshDelayMs: 0,
        refreshTimer: null,
        activeRequestSeq: 0,
        sceneRevision: 0,
      } satisfies SceneEntry)
    if (existing) {
      entry.request = preferSceneRequest(existing.request, request)
    }
    entry.listeners.add(listener)
    this.entries.set(key, entry)
    if (entry.scenes === null) {
      this.publishImmediateScenes(entry, { countRefresh: false })
      this.queueRefresh(entry)
    } else if (!existing && retained?.source === 'immediate') {
      this.queueRefresh(entry)
    }
    return () => {
      entry.listeners.delete(listener)
      this.releaseEntryIfInactive(entry)
    }
  }

  dropSheets(sheetNames: readonly string[]): void {
    if (sheetNames.length === 0) {
      return
    }
    const removed = new Set(sheetNames)
    for (const entry of this.entries.values()) {
      if (removed.has(entry.request.sheetName)) {
        this.dropEntry(entry)
      }
    }
    for (const [key, scenes] of this.retainedScenes) {
      if (removed.has(scenes.request.sheetName)) {
        this.retainedScenes.delete(key)
      }
    }
  }

  noteViewportPatch(patch: ViewportPatch, applied?: ResidentScenePatchDamage): void {
    this.refreshMatchingSceneEntries(
      (request) => residentPaneSceneRequestNeedsRefresh(request, patch, applied),
      (scenes) => !residentScenesCoverViewportPatch(scenes, patch, applied),
    )
  }

  noteCellDamage(sheetName: string, row: number, col: number): void {
    this.refreshMatchingEntries(
      (request) =>
        request.sheetName === sheetName &&
        row >= request.residentViewport.rowStart &&
        row <= request.residentViewport.rowEnd &&
        col >= request.residentViewport.colStart &&
        col <= request.residentViewport.colEnd,
    )
  }

  private refreshMatchingEntries(matches: (request: WorkbookPaneSceneRequest) => boolean): void {
    this.refreshMatchingSceneEntries(matches, () => true)
  }

  private refreshMatchingSceneEntries(
    matches: (request: WorkbookPaneSceneRequest) => boolean,
    shouldRefresh: (scenes: readonly WorkbookPaneScenePacket[] | null) => boolean,
  ): void {
    this.deleteMatchingRetainedScenes((retained) => matches(retained.request) && shouldRefresh(retained.scenes))
    for (const entry of this.entries.values()) {
      if (!matches(entry.request)) {
        continue
      }
      if (!shouldRefresh(entry.coverageScenes ?? entry.scenes)) {
        continue
      }
      entry.sceneRevision += 1
      const publishedImmediateScenes = this.publishImmediateScenes(entry)
      if (!publishedImmediateScenes) {
        this.queueRefresh(entry, BACKGROUND_WORKER_SCENE_REFRESH_DELAY_MS)
      }
    }
  }

  private publishImmediateScenes(entry: SceneEntry, options: { readonly countRefresh?: boolean } = {}): boolean {
    const buildImmediateResidentPaneScenes = this.options.buildImmediateResidentPaneScenes
    if (!buildImmediateResidentPaneScenes || entry.listeners.size === 0) {
      return false
    }

    const generation = this.allocateImmediateSceneGeneration(entry)
    const request = withRequestMetadata(entry.request, entry.activeRequestSeq, entry.sceneRevision)
    let next: readonly WorkbookPaneScenePacket[] | null
    try {
      next = buildImmediateResidentPaneScenes(request, generation)
    } catch {
      this.noteScenePacketRejected('immediate-scene-build-failed')
      return false
    }
    if (next === null) {
      return false
    }
    if (!isResidentPaneScenePacketArray(next)) {
      this.noteScenePacketRejected('invalid-immediate-scene-payload')
      return false
    }
    if (next.length === 0) {
      return false
    }

    const sameDrawSemantics = sameResidentSceneDrawSemantics(entry.scenes, next)
    this.retainScenes(entry.request, next, 'immediate')
    if (sameDrawSemantics) {
      entry.coverageScenes = next
      return true
    }
    entry.scenes = next
    entry.coverageScenes = next
    if (options.countRefresh ?? true) {
      getWorkbookScrollPerfCollector()?.noteScenePacketRefresh(next.length)
    }
    entry.listeners.forEach((listener) => listener())
    return true
  }

  private allocateImmediateSceneGeneration(entry: SceneEntry): number {
    const currentGeneration = entry.scenes?.reduce((max, scene) => Math.max(max, scene.generation), 0) ?? 0
    this.nextImmediateSceneGeneration = Math.max(this.nextImmediateSceneGeneration, currentGeneration) + 1
    return this.nextImmediateSceneGeneration
  }

  private queueRefresh(entry: SceneEntry, delayMs = 0): void {
    entry.refreshDelayMs = Math.max(entry.refreshDelayMs, delayMs)
    entry.refreshQueued = true
    if (entry.inFlight) {
      return
    }
    this.scheduleRefresh(entry)
  }

  private scheduleRefresh(entry: SceneEntry): void {
    if (entry.refreshTimer) {
      clearTimeout(entry.refreshTimer)
      entry.refreshTimer = null
    }
    const delayMs = entry.refreshDelayMs
    const flush = () => {
      entry.refreshTimer = null
      if (!this.isEntryCurrent(entry) || !entry.refreshQueued || entry.inFlight || entry.listeners.size === 0) {
        return
      }
      entry.refreshQueued = false
      entry.refreshDelayMs = 0
      void this.refreshEntry(entry)
    }
    if (delayMs > 0) {
      entry.refreshTimer = setTimeout(flush, delayMs)
      return
    }
    queueMicrotask(flush)
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
        const requestSeq = ++this.nextRequestSeq
        entry.activeRequestSeq = requestSeq
        const request = withRequestMetadata(entry.request, requestSeq, entry.sceneRevision)
        const isIncrementalRefresh = entry.scenes !== null
        const next = await client.invoke('getResidentPaneScenes', request)
        if (!this.isRefreshCurrent(entry, requestSeq)) {
          return
        }
        if (!isResidentPaneScenePacketArray(next)) {
          this.noteScenePacketRejected('invalid-scene-payload')
          return
        }
        if (!this.isRefreshCurrent(entry, requestSeq)) {
          return
        }
        const packetRejectionReason = validateScenePacketResponseForRequest(next, request)
        if (packetRejectionReason !== null) {
          this.noteScenePacketRejected(packetRejectionReason)
          return
        }
        if (sameResidentSceneDrawSemantics(entry.scenes, next)) {
          entry.coverageScenes = next
          this.retainScenes(entry.request, next, 'worker')
          return
        }
        entry.scenes = next
        entry.coverageScenes = next
        this.retainScenes(entry.request, next, 'worker')
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
      this.releaseEntryIfInactive(entry)
    }
  }

  private peekRetainedScenes(request: WorkbookPaneSceneRequest): RetainedSceneEntry | null {
    const key = buildRetainedRequestKey(request)
    const retained = this.retainedScenes.get(key)
    if (!retained) {
      return null
    }
    const next = {
      ...retained,
      lastUsedSeq: ++this.retainedSceneSeq,
    }
    this.retainedScenes.set(key, next)
    return next
  }

  private buildAndRetainImmediateScenes(request: WorkbookPaneSceneRequest): readonly WorkbookPaneScenePacket[] | null {
    return this.buildAndRetainImmediateSceneEntry(request)?.scenes ?? null
  }

  private buildAndRetainImmediateSceneEntry(request: WorkbookPaneSceneRequest): RetainedSceneEntry | null {
    const scenes = this.buildImmediateSceneSnapshot(request, ++this.nextImmediateSceneGeneration, 0, 0)
    if (!scenes) {
      return null
    }
    return this.retainScenes(request, scenes, 'immediate')
  }

  private buildImmediateSceneSnapshot(
    request: WorkbookPaneSceneRequest,
    generation: number,
    requestSeq: number,
    sceneRevision: number,
  ): readonly WorkbookPaneScenePacket[] | null {
    const buildImmediateResidentPaneScenes = this.options.buildImmediateResidentPaneScenes
    if (!buildImmediateResidentPaneScenes) {
      return null
    }
    const sceneRequest = withRequestMetadata(request, requestSeq, sceneRevision)
    let next: readonly WorkbookPaneScenePacket[] | null
    try {
      next = buildImmediateResidentPaneScenes(sceneRequest, generation)
    } catch {
      this.noteScenePacketRejected('immediate-scene-build-failed')
      return null
    }
    if (next === null) {
      return null
    }
    if (!isResidentPaneScenePacketArray(next)) {
      this.noteScenePacketRejected('invalid-immediate-scene-payload')
      return null
    }
    return next.length === 0 ? null : next
  }

  private retainScenes(
    request: WorkbookPaneSceneRequest,
    scenes: readonly WorkbookPaneScenePacket[],
    source: RetainedSceneEntry['source'],
  ): RetainedSceneEntry {
    const entry = {
      lastUsedSeq: ++this.retainedSceneSeq,
      request,
      scenes,
      source,
    }
    this.retainedScenes.set(buildRetainedRequestKey(request), {
      ...entry,
    })
    this.evictRetainedScenes()
    return entry
  }

  private evictRetainedScenes(): void {
    if (this.retainedScenes.size <= MAX_RETAINED_RESIDENT_SCENE_REQUESTS) {
      return
    }
    const retained = [...this.retainedScenes.entries()].toSorted((left, right) => left[1].lastUsedSeq - right[1].lastUsedSeq)
    for (const [key] of retained) {
      if (this.retainedScenes.size <= MAX_RETAINED_RESIDENT_SCENE_REQUESTS) {
        return
      }
      this.retainedScenes.delete(key)
    }
  }

  private deleteMatchingRetainedScenes(matches: (retained: RetainedSceneEntry) => boolean): void {
    for (const [key, retained] of this.retainedScenes) {
      if (matches(retained)) {
        this.retainedScenes.delete(key)
      }
    }
  }

  private isEntryCurrent(entry: SceneEntry): boolean {
    return this.entries.get(entry.key) === entry
  }

  private isRefreshCurrent(entry: SceneEntry, requestSeq: number): boolean {
    return this.isEntryCurrent(entry) && entry.activeRequestSeq === requestSeq && !entry.refreshQueued && entry.listeners.size > 0
  }

  private clearEntryScheduledWork(entry: SceneEntry): void {
    if (entry.refreshTimer) {
      clearTimeout(entry.refreshTimer)
      entry.refreshTimer = null
    }
    entry.refreshQueued = false
    entry.refreshDelayMs = 0
    entry.activeRequestSeq += 1
  }

  private dropEntry(entry: SceneEntry): void {
    this.clearEntryScheduledWork(entry)
    entry.listeners.clear()
    entry.scenes = null
    entry.coverageScenes = null
    if (this.entries.get(entry.key) === entry) {
      this.entries.delete(entry.key)
    }
  }

  private releaseEntryIfInactive(entry: SceneEntry): void {
    if (entry.listeners.size > 0) {
      return
    }
    this.clearEntryScheduledWork(entry)
    if (this.entries.get(entry.key) === entry) {
      this.entries.delete(entry.key)
    }
  }

  private noteScenePacketRejected(reason: string): void {
    this.lastRejectedPacketReason = reason
    getWorkbookScrollPerfCollector()?.noteScenePacketRejected(reason)
  }
}
