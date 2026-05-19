import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridRenderRevisionSnapshot } from '../grid-engine.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { DynamicGridOverlayBatchV3 } from './dynamic-overlay-batch.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveGridTextTileRevisionKeyV3 } from './typegpu-tile-resource-revisions.js'
import { WorkbookPaneRendererRuntimeV3, type WorkbookPaneFrameResultV3 } from './workbook-pane-renderer-runtime.js'
import {
  EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3,
  WorkbookPaneSurfaceRuntimeV3,
  type WorkbookPaneSurfaceBackendStatusV3,
  type WorkbookPaneSurfaceSnapshotV3,
} from './workbook-pane-surface-runtime.js'

export interface WorkbookPaneRendererHostPropsV3 {
  readonly active: boolean
  readonly cameraStore: GridCameraStore | null
  readonly drawText: boolean
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly host: HTMLDivElement | null
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly overlayBuilder: ((geometry: GridGeometrySnapshot) => DynamicGridOverlayBatchV3 | null | undefined) | null
  readonly preloadTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly renderRevisionSnapshot: GridRenderRevisionSnapshot | null
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

export interface WorkbookPaneRendererHostRuntimeOptionsV3 {
  readonly rendererRuntime?: WorkbookPaneRendererRuntimeV3 | undefined
  readonly surfaceRuntime?: WorkbookPaneSurfaceRuntimeV3 | undefined
}

export type WorkbookPaneFrameProofStatusV3 = 'idle' | 'pending' | 'presented'

const EMPTY_HOST_PROPS: WorkbookPaneRendererHostPropsV3 = Object.freeze({
  active: false,
  cameraStore: null,
  drawText: true,
  geometry: null,
  headerPanes: [],
  host: null,
  overlay: null,
  overlayBuilder: null,
  preloadTilePanes: [],
  renderRevisionSnapshot: null,
  scrollTransformStore: null,
  tilePanes: [],
})

export class WorkbookPaneRendererHostRuntimeV3 {
  private canvas: HTMLCanvasElement | null = null
  private disposed = false
  private readonly backendStatusListeners = new Set<() => void>()
  private readonly frameProofListeners = new Set<() => void>()
  private frameProofSignature = ''
  private frameProofStatus: WorkbookPaneFrameProofStatusV3 = 'idle'
  private hasPresentedFrame = false
  private presentedFrameProofSignature = ''
  private props: WorkbookPaneRendererHostPropsV3 = EMPTY_HOST_PROPS
  private readonly rendererRuntime: WorkbookPaneRendererRuntimeV3
  private surfaceBackendStatus: WorkbookPaneSurfaceBackendStatusV3
  private surfaceSnapshot: WorkbookPaneSurfaceSnapshotV3 = EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3
  private readonly surfaceRuntime: WorkbookPaneSurfaceRuntimeV3
  private readonly unsubscribeSurface: () => void

  constructor(options: WorkbookPaneRendererHostRuntimeOptionsV3 = {}) {
    this.rendererRuntime = options.rendererRuntime ?? new WorkbookPaneRendererRuntimeV3()
    this.rendererRuntime.setFrameResultListener((result) => this.handleFrameResult(result))
    this.surfaceRuntime = options.surfaceRuntime ?? new WorkbookPaneSurfaceRuntimeV3()
    this.surfaceBackendStatus = this.surfaceRuntime.getSnapshot().backendStatus
    this.unsubscribeSurface = this.surfaceRuntime.subscribe((snapshot) => {
      this.surfaceSnapshot = snapshot
      if (this.surfaceBackendStatus !== snapshot.backendStatus) {
        this.surfaceBackendStatus = snapshot.backendStatus
        this.emitBackendStatus()
      }
      if (snapshot.backendStatus !== 'ready') {
        this.setFrameProofStatus(this.frameProofSignature ? 'pending' : 'idle')
      }
      this.applyRendererState()
      this.requestRenderDraw()
    })
  }

  readonly getBackendStatusSnapshot = (): WorkbookPaneSurfaceBackendStatusV3 => this.surfaceBackendStatus
  readonly getFrameProofSignatureSnapshot = (): string => this.frameProofSignature
  readonly getFrameProofStatusSnapshot = (): WorkbookPaneFrameProofStatusV3 => this.frameProofStatus
  readonly getHasPresentedFrameSnapshot = (): boolean => this.hasPresentedFrame
  readonly getPresentedFrameProofSignatureSnapshot = (): string => this.presentedFrameProofSignature

  readonly subscribeBackendStatus = (listener: () => void): (() => void) => {
    if (this.disposed) {
      return () => {}
    }
    this.backendStatusListeners.add(listener)
    return () => {
      this.backendStatusListeners.delete(listener)
    }
  }

  readonly subscribeFrameProofStatus = (listener: () => void): (() => void) => {
    if (this.disposed) {
      return () => {}
    }
    this.frameProofListeners.add(listener)
    return () => {
      this.frameProofListeners.delete(listener)
    }
  }

  updateProps(props: WorkbookPaneRendererHostPropsV3): void {
    if (this.disposed) {
      return
    }
    this.syncFrameProofSignature(props)
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
    this.frameProofListeners.clear()
    this.rendererRuntime.setFrameResultListener(null)
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
      drawText: this.props.drawText,
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

  private emitFrameProofStatus(): void {
    this.frameProofListeners.forEach((listener) => listener())
  }

  private handleFrameResult(result: WorkbookPaneFrameResultV3): void {
    const signature = this.frameProofSignature
    if (!result.submitted || !signature) {
      return
    }
    this.setPresentedFrameProofSignature(signature)
    this.setHasPresentedFrame(true)
    this.setFrameProofStatus('presented')
  }

  private setFrameProofStatus(status: WorkbookPaneFrameProofStatusV3): void {
    if (this.frameProofStatus === status) {
      return
    }
    this.frameProofStatus = status
    this.emitFrameProofStatus()
  }

  private setHasPresentedFrame(value: boolean): void {
    if (this.hasPresentedFrame === value) {
      return
    }
    this.hasPresentedFrame = value
    this.emitFrameProofStatus()
  }

  private setPresentedFrameProofSignature(signature: string): void {
    if (this.presentedFrameProofSignature === signature) {
      return
    }
    this.presentedFrameProofSignature = signature
    this.emitFrameProofStatus()
  }

  private syncFrameProofSignature(props: WorkbookPaneRendererHostPropsV3): void {
    const signature = resolveWorkbookPaneFrameProofSignatureV3(props)
    if (this.frameProofSignature === signature) {
      return
    }
    this.frameProofSignature = signature
    if (!signature) {
      this.setPresentedFrameProofSignature('')
      this.setHasPresentedFrame(false)
      this.setFrameProofStatus('idle')
      return
    }
    const hasPresentedCurrentSignature = this.presentedFrameProofSignature === signature
    this.setHasPresentedFrame(hasPresentedCurrentSignature)
    this.setFrameProofStatus(hasPresentedCurrentSignature ? 'presented' : 'pending')
  }

  private syncCanvasTarget(): void {
    if (this.disposed) {
      return
    }
    this.surfaceRuntime.setCanvas(this.props.active && this.props.host ? this.canvas : null)
  }
}

export function resolveWorkbookPaneFrameProofSignatureV3(props: {
  readonly drawText?: boolean | undefined
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly overlay: DynamicGridOverlayBatchV3 | null
  readonly renderRevisionSnapshot?: GridRenderRevisionSnapshot | null | undefined
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}): string {
  const textOwnershipSignature = `drawText:${props.drawText === false ? 'gpu-text-off' : 'gpu-text-on'}`
  const renderRevisionSignature = props.renderRevisionSnapshot
    ? [
        props.renderRevisionSnapshot.authoritativeRevision ?? 'none',
        props.renderRevisionSnapshot.localRevision ?? 'none',
        props.renderRevisionSnapshot.projectedRevision,
        props.renderRevisionSnapshot.tileSceneCameraSeq ?? 'none',
        props.renderRevisionSnapshot.tileSceneRevision ?? 'none',
      ].join(':')
    : ''
  const tileSignature = props.tilePanes
    .map((pane) => {
      const tile = pane.tile
      return [
        pane.paneId,
        pane.generation,
        tile.tileId,
        tile.textCount,
        tile.textSignature ?? resolveGridTextTileRevisionKeyV3(tile).textSignature,
        tile.rectCount,
        tile.rectSignature ?? '',
        tile.version.axisX,
        tile.version.axisY,
        tile.version.freeze,
        tile.version.styles,
        tile.version.text,
        tile.version.values,
      ].join(':')
    })
    .join('|')
  const headerSignature = props.headerPanes
    .map((pane) => [pane.paneId, pane.rectSignature, pane.textSignature, pane.rectCount, pane.textCount].join(':'))
    .join('|')
  const overlaySignature = props.overlay ? `${props.overlay.seq}:${props.overlay.rectCount}` : ''
  return [textOwnershipSignature, tileSignature, headerSignature, overlaySignature, renderRevisionSignature].filter(Boolean).join('#')
}
