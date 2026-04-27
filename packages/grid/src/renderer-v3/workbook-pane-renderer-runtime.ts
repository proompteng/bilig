import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { GridDrawSchedulerV3 } from './draw-scheduler.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { drawWorkbookTypeGpuTileFrameV3, type WorkbookTypeGpuBackendV3 } from './typegpu-workbook-backend-v3.js'

export interface TypeGpuSurfaceSizeV3 {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

export interface WorkbookPaneRendererRuntimeStateV3 {
  readonly active: boolean
  readonly backend: unknown
  readonly cameraStore: GridCameraStore | null
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null
  readonly preloadTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly surface: TypeGpuSurfaceSizeV3
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly webGpuReady: boolean
}

export interface WorkbookPaneFrameInputV3 {
  readonly backend: unknown
  readonly headerPanes?: readonly GridHeaderPaneState[] | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
  readonly preloadTilePanes?: readonly WorkbookRenderTilePaneState[] | undefined
  readonly overlay?: DynamicGridOverlayBatchV3 | null | undefined
  readonly syncPreloadPanes?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuSurfaceSizeV3
}

export type WorkbookPaneFrameDrawerV3 = (input: WorkbookPaneFrameInputV3) => void

const EMPTY_SURFACE_SIZE: TypeGpuSurfaceSizeV3 = Object.freeze({
  dpr: 1,
  height: 0,
  pixelHeight: 0,
  pixelWidth: 0,
  width: 0,
})

const EMPTY_RUNTIME_STATE: WorkbookPaneRendererRuntimeStateV3 = Object.freeze({
  active: false,
  backend: null,
  cameraStore: null,
  geometry: null,
  headerPanes: [],
  overlay: null,
  overlayBuilder: null,
  preloadTilePanes: [],
  scrollTransformStore: null,
  surface: EMPTY_SURFACE_SIZE,
  tilePanes: [],
  webGpuReady: false,
})

function isWorkbookTypeGpuBackendV3(value: unknown): value is WorkbookTypeGpuBackendV3 {
  return (
    typeof value === 'object' &&
    value !== null &&
    'artifacts' in value &&
    'atlas' in value &&
    'layerResources' in value &&
    'surfaceState' in value &&
    'tileResources' in value &&
    'tileResidency' in value
  )
}

function drawWorkbookPaneFrameV3(input: WorkbookPaneFrameInputV3): void {
  if (!isWorkbookTypeGpuBackendV3(input.backend)) {
    return
  }
  drawWorkbookTypeGpuTileFrameV3({
    backend: input.backend,
    headerPanes: input.headerPanes,
    overlay: input.overlay,
    preloadTilePanes: input.preloadTilePanes,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
    syncPreloadPanes: input.syncPreloadPanes,
    tilePanes: input.tilePanes,
  })
}

export function resolveTypeGpuV3DrawScrollSnapshot(input: {
  readonly fallback: WorkbookGridScrollSnapshot
  readonly geometry: GridGeometrySnapshot | null
  readonly panes: readonly WorkbookRenderTilePaneState[]
}): WorkbookGridScrollSnapshot {
  const bodyPane = input.panes.find((pane) => pane.paneId === 'body')
  if (!input.geometry || !bodyPane) {
    return input.fallback
  }

  const bodyWorldX = input.geometry.camera.frozenWidth + (input.fallback.scrollLeft ?? input.geometry.camera.bodyScrollX)
  const bodyWorldY = input.geometry.camera.frozenHeight + (input.fallback.scrollTop ?? input.geometry.camera.bodyScrollY)
  return {
    ...input.fallback,
    renderTx: bodyWorldX - input.geometry.columns.offsetOf(bodyPane.viewport.colStart),
    renderTy: bodyWorldY - input.geometry.rows.offsetOf(bodyPane.viewport.rowStart),
  }
}

export class WorkbookPaneRendererRuntimeV3 {
  private cameraStoreUnsubscribe: (() => void) | null = null
  private scrollStoreUnsubscribe: (() => void) | null = null
  private state: WorkbookPaneRendererRuntimeStateV3 = EMPTY_RUNTIME_STATE
  private subscribedCameraStore: GridCameraStore | null = null
  private subscribedScrollStore: WorkbookGridScrollStore | null = null

  constructor(
    private readonly drawFrame: WorkbookPaneFrameDrawerV3 = drawWorkbookPaneFrameV3,
    private readonly scheduler = new GridDrawSchedulerV3(),
  ) {}

  updateState(state: Partial<WorkbookPaneRendererRuntimeStateV3>): void {
    this.state = {
      ...this.state,
      ...state,
    }
    this.syncStoreSubscriptions()
  }

  requestDraw(): void {
    this.scheduler.requestDraw(() => this.drawNow())
  }

  noteInputSignalAndRequestDraw(): void {
    this.scheduler.noteInputSignal()
    this.requestDraw()
  }

  drawNow(): void {
    const state = this.state
    if (
      !state.active ||
      !state.webGpuReady ||
      state.backend === null ||
      state.backend === undefined ||
      state.surface.width <= 0 ||
      state.surface.height <= 0
    ) {
      return
    }

    const latestGeometry = state.cameraStore?.getSnapshot() ?? state.geometry
    const frameDecision = this.scheduler.resolveFrame({
      camera: latestGeometry?.camera ?? null,
      requestIdlePreloadDraw: () => this.requestDraw(),
    })
    const overlayBatch = state.overlayBuilder && latestGeometry ? state.overlayBuilder(latestGeometry) : state.overlay

    this.drawFrame({
      backend: state.backend,
      headerPanes: state.headerPanes,
      overlay: overlayBatch ?? null,
      preloadTilePanes: state.preloadTilePanes,
      scrollSnapshot: resolveTypeGpuV3DrawScrollSnapshot({
        fallback: state.scrollTransformStore?.getSnapshot() ?? { tx: 0, ty: 0 },
        geometry: latestGeometry,
        panes: state.tilePanes,
      }),
      surface: state.surface,
      syncPreloadPanes: frameDecision.syncPreloadPanes,
      tilePanes: state.tilePanes,
    })
  }

  dispose(): void {
    this.clearStoreSubscriptions()
    this.scheduler.cancel()
    this.state = EMPTY_RUNTIME_STATE
  }

  private syncStoreSubscriptions(): void {
    const nextCameraStore = this.state.active ? this.state.cameraStore : null
    if (this.subscribedCameraStore !== nextCameraStore) {
      this.cameraStoreUnsubscribe?.()
      this.cameraStoreUnsubscribe = null
      this.subscribedCameraStore = nextCameraStore
      if (nextCameraStore) {
        this.cameraStoreUnsubscribe = nextCameraStore.subscribe(() => this.noteInputSignalAndRequestDraw())
      }
    }

    const nextScrollStore = this.state.active ? this.state.scrollTransformStore : null
    if (this.subscribedScrollStore !== nextScrollStore) {
      this.scrollStoreUnsubscribe?.()
      this.scrollStoreUnsubscribe = null
      this.subscribedScrollStore = nextScrollStore
      if (nextScrollStore) {
        this.scrollStoreUnsubscribe = nextScrollStore.subscribe(() => this.noteInputSignalAndRequestDraw())
      }
    }
  }

  private clearStoreSubscriptions(): void {
    this.cameraStoreUnsubscribe?.()
    this.scrollStoreUnsubscribe?.()
    this.cameraStoreUnsubscribe = null
    this.scrollStoreUnsubscribe = null
    this.subscribedCameraStore = null
    this.subscribedScrollStore = null
  }
}
