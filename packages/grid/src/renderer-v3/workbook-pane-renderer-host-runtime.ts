import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { WorkbookPaneRendererRuntimeV3 } from './workbook-pane-renderer-runtime.js'
import {
  EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3,
  WorkbookPaneSurfaceRuntimeV3,
  type WorkbookPaneSurfaceBackendStatusV3,
  type WorkbookPaneSurfaceSnapshotV3,
} from './workbook-pane-surface-runtime.js'

export interface WorkbookPaneRendererHostPropsV3 {
  readonly active: boolean
  readonly cameraStore: GridCameraStore | null
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly host: HTMLDivElement | null
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null
  readonly preloadTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

export interface WorkbookPaneRendererHostRuntimeOptionsV3 {
  readonly rendererRuntime?: WorkbookPaneRendererRuntimeV3 | undefined
  readonly surfaceRuntime?: WorkbookPaneSurfaceRuntimeV3 | undefined
}

const EMPTY_HOST_PROPS: WorkbookPaneRendererHostPropsV3 = Object.freeze({
  active: false,
  cameraStore: null,
  geometry: null,
  headerPanes: [],
  host: null,
  overlay: null,
  overlayBuilder: null,
  preloadTilePanes: [],
  scrollTransformStore: null,
  tilePanes: [],
})

export class WorkbookPaneRendererHostRuntimeV3 {
  private canvas: HTMLCanvasElement | null = null
  private disposed = false
  private readonly backendStatusListeners = new Set<() => void>()
  private props: WorkbookPaneRendererHostPropsV3 = EMPTY_HOST_PROPS
  private readonly rendererRuntime: WorkbookPaneRendererRuntimeV3
  private surfaceBackendStatus: WorkbookPaneSurfaceBackendStatusV3
  private surfaceSnapshot: WorkbookPaneSurfaceSnapshotV3 = EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3
  private readonly surfaceRuntime: WorkbookPaneSurfaceRuntimeV3
  private readonly unsubscribeSurface: () => void

  constructor(options: WorkbookPaneRendererHostRuntimeOptionsV3 = {}) {
    this.rendererRuntime = options.rendererRuntime ?? new WorkbookPaneRendererRuntimeV3()
    this.surfaceRuntime = options.surfaceRuntime ?? new WorkbookPaneSurfaceRuntimeV3()
    this.surfaceBackendStatus = this.surfaceRuntime.getSnapshot().backendStatus
    this.unsubscribeSurface = this.surfaceRuntime.subscribe((snapshot) => {
      this.surfaceSnapshot = snapshot
      if (this.surfaceBackendStatus !== snapshot.backendStatus) {
        this.surfaceBackendStatus = snapshot.backendStatus
        this.emitBackendStatus()
      }
      this.applyRendererState()
      this.requestRenderDraw()
    })
  }

  readonly getBackendStatusSnapshot = (): WorkbookPaneSurfaceBackendStatusV3 => this.surfaceBackendStatus

  readonly subscribeBackendStatus = (listener: () => void): (() => void) => {
    if (this.disposed) {
      return () => {}
    }
    this.backendStatusListeners.add(listener)
    return () => {
      this.backendStatusListeners.delete(listener)
    }
  }

  updateProps(props: WorkbookPaneRendererHostPropsV3): void {
    if (this.disposed) {
      return
    }
    this.props = props
    this.surfaceRuntime.setHost(props.host)
    this.surfaceRuntime.setActive(props.active)
    this.syncCanvasTarget()
    this.applyRendererState()
    this.requestRenderDraw()
  }

  setCanvas(canvas: HTMLCanvasElement | null): void {
    if (this.canvas === canvas) {
      return
    }
    this.canvas = canvas
    this.syncCanvasTarget()
  }

  dispose(): void {
    if (this.disposed) {
      return
    }
    this.disposed = true
    this.unsubscribeSurface()
    this.backendStatusListeners.clear()
    const canvas = this.canvas
    this.canvas = null
    this.surfaceRuntime.dispose()
    this.rendererRuntime.dispose()
    if (canvas) {
      canvas.width = 0
      canvas.height = 0
    }
  }

  private applyRendererState(): void {
    this.rendererRuntime.updateState({
      active: this.props.active,
      backend: this.surfaceSnapshot.backend,
      cameraStore: this.props.cameraStore,
      geometry: this.props.geometry,
      headerPanes: this.props.headerPanes,
      overlay: this.props.overlay,
      overlayBuilder: this.props.overlayBuilder,
      preloadTilePanes: this.props.preloadTilePanes,
      scrollTransformStore: this.props.scrollTransformStore,
      surface: this.surfaceSnapshot.surface,
      tilePanes: this.props.tilePanes,
      webGpuReady: this.surfaceSnapshot.webGpuReady,
    })
  }

  private requestRenderDraw(): void {
    this.rendererRuntime.requestDraw()
  }

  private emitBackendStatus(): void {
    this.backendStatusListeners.forEach((listener) => listener())
  }

  private syncCanvasTarget(): void {
    if (this.disposed) {
      return
    }
    this.surfaceRuntime.setCanvas(this.props.active && this.props.host ? this.canvas : null)
  }
}
