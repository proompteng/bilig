import {
  createWorkbookTypeGpuBackendV3,
  destroyWorkbookTypeGpuBackendV3,
  syncWorkbookTypeGpuSurfaceV3,
  type WorkbookTypeGpuBackendV3,
} from './typegpu-workbook-backend-v3.js'
import type { TypeGpuSurfaceSizeV3 } from './workbook-pane-renderer-runtime.js'

export interface WorkbookPaneSurfaceSnapshotV3 {
  readonly backend: object | null
  readonly surface: TypeGpuSurfaceSizeV3
  readonly webGpuReady: boolean
}

export interface WorkbookPaneSurfaceRuntimeResizeObserverV3 {
  observe(target: Element): void
  disconnect(): void
}

export interface WorkbookPaneSurfaceRuntimeOptionsV3 {
  readonly createBackend?: ((canvas: HTMLCanvasElement) => Promise<object | null>) | undefined
  readonly destroyBackend?: ((backend: object) => void) | undefined
  readonly syncSurface?:
    | ((input: { readonly backend: object; readonly canvas: HTMLCanvasElement; readonly size: TypeGpuSurfaceSizeV3 }) => void)
    | undefined
  readonly getDevicePixelRatio?: (() => number) | undefined
  readonly createResizeObserver?: ((listener: ResizeObserverCallback) => WorkbookPaneSurfaceRuntimeResizeObserverV3 | null) | undefined
}

export const EMPTY_TYPEGPU_SURFACE_SIZE_V3: TypeGpuSurfaceSizeV3 = Object.freeze({
  dpr: 1,
  height: 0,
  pixelHeight: 0,
  pixelWidth: 0,
  width: 0,
})

export const EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3: WorkbookPaneSurfaceSnapshotV3 = Object.freeze({
  backend: null,
  surface: EMPTY_TYPEGPU_SURFACE_SIZE_V3,
  webGpuReady: false,
})

export function resolveWorkbookPaneSurfaceSizeV3(input: {
  readonly host: Pick<HTMLElement, 'clientHeight' | 'clientWidth'>
  readonly dpr?: number | undefined
}): TypeGpuSurfaceSizeV3 {
  const width = Math.max(0, Math.floor(input.host.clientWidth))
  const height = Math.max(0, Math.floor(input.host.clientHeight))
  const dpr = Math.max(1, input.dpr ?? defaultDevicePixelRatio())
  return {
    dpr,
    height,
    pixelHeight: Math.max(1, Math.floor(height * dpr)),
    pixelWidth: Math.max(1, Math.floor(width * dpr)),
    width,
  }
}

function defaultDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1
}

function createDefaultResizeObserver(listener: ResizeObserverCallback): WorkbookPaneSurfaceRuntimeResizeObserverV3 | null {
  if (typeof ResizeObserver === 'undefined') {
    return null
  }
  return new ResizeObserver(listener)
}

function createDefaultBackend(canvas: HTMLCanvasElement): Promise<object | null> {
  return createWorkbookTypeGpuBackendV3(canvas)
}

function isWorkbookTypeGpuBackendV3(value: object): value is WorkbookTypeGpuBackendV3 {
  return (
    'artifacts' in value &&
    'atlas' in value &&
    'layerResources' in value &&
    'surfaceState' in value &&
    'tileResources' in value &&
    'tileResidency' in value
  )
}

function destroyDefaultBackend(backend: object): void {
  if (!isWorkbookTypeGpuBackendV3(backend)) {
    return
  }
  destroyWorkbookTypeGpuBackendV3(backend)
}

function syncDefaultSurface(input: {
  readonly backend: object
  readonly canvas: HTMLCanvasElement
  readonly size: TypeGpuSurfaceSizeV3
}): void {
  if (!isWorkbookTypeGpuBackendV3(input.backend)) {
    return
  }
  syncWorkbookTypeGpuSurfaceV3({
    backend: input.backend,
    canvas: input.canvas,
    size: input.size,
  })
}

export class WorkbookPaneSurfaceRuntimeV3 {
  private active = false
  private backend: object | null = null
  private canvas: HTMLCanvasElement | null = null
  private host: HTMLElement | null = null
  private initToken = 0
  private listener: ((snapshot: WorkbookPaneSurfaceSnapshotV3) => void) | null = null
  private resizeObserver: WorkbookPaneSurfaceRuntimeResizeObserverV3 | null = null
  private snapshot: WorkbookPaneSurfaceSnapshotV3 = EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3

  private readonly createBackend: (canvas: HTMLCanvasElement) => Promise<object | null>
  private readonly destroyBackend: (backend: object) => void
  private readonly syncSurface: (input: {
    readonly backend: object
    readonly canvas: HTMLCanvasElement
    readonly size: TypeGpuSurfaceSizeV3
  }) => void
  private readonly getDevicePixelRatio: () => number
  private readonly createResizeObserver: (listener: ResizeObserverCallback) => WorkbookPaneSurfaceRuntimeResizeObserverV3 | null

  constructor(options: WorkbookPaneSurfaceRuntimeOptionsV3 = {}) {
    this.createBackend = options.createBackend ?? createDefaultBackend
    this.destroyBackend = options.destroyBackend ?? destroyDefaultBackend
    this.syncSurface = options.syncSurface ?? syncDefaultSurface
    this.getDevicePixelRatio = options.getDevicePixelRatio ?? defaultDevicePixelRatio
    this.createResizeObserver = options.createResizeObserver ?? createDefaultResizeObserver
  }

  getSnapshot(): WorkbookPaneSurfaceSnapshotV3 {
    return this.snapshot
  }

  subscribe(listener: (snapshot: WorkbookPaneSurfaceSnapshotV3) => void): () => void {
    this.listener = listener
    listener(this.snapshot)
    return () => {
      if (this.listener === listener) {
        this.listener = null
      }
    }
  }

  setActive(active: boolean): void {
    if (this.active === active) {
      return
    }
    this.active = active
    if (!active) {
      this.destroyCurrentBackend()
      return
    }
    void this.ensureBackend()
  }

  setCanvas(canvas: HTMLCanvasElement | null): void {
    if (this.canvas === canvas) {
      return
    }
    if (this.backend) {
      this.destroyCurrentBackend()
    }
    this.canvas = canvas
    void this.ensureBackend()
  }

  setHost(host: HTMLElement | null): void {
    if (this.host === host) {
      return
    }
    this.resizeObserver?.disconnect()
    this.resizeObserver = null
    this.host = host
    this.refreshSurface()
    if (!host) {
      return
    }
    const observer = this.createResizeObserver(() => {
      this.refreshSurface()
    })
    if (!observer) {
      return
    }
    observer.observe(host)
    this.resizeObserver = observer
  }

  dispose(): void {
    this.resizeObserver?.disconnect()
    this.resizeObserver = null
    this.host = null
    this.canvas = null
    this.active = false
    this.listener = null
    this.destroyCurrentBackend()
    this.updateSnapshot(EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3)
  }

  private refreshSurface(): void {
    const surface = this.host
      ? resolveWorkbookPaneSurfaceSizeV3({
          dpr: this.getDevicePixelRatio(),
          host: this.host,
        })
      : EMPTY_TYPEGPU_SURFACE_SIZE_V3
    this.updateSnapshot({
      backend: this.backend,
      surface,
      webGpuReady: this.backend !== null,
    })
  }

  private async ensureBackend(): Promise<void> {
    const canvas = this.canvas
    if (!this.active || !canvas || this.backend) {
      return
    }
    const token = ++this.initToken
    const backend = await this.createBackend(canvas)
    if (token !== this.initToken || !this.active || this.canvas !== canvas) {
      if (backend) {
        this.destroyBackend(backend)
      }
      return
    }
    this.backend = backend
    this.updateSnapshot({
      backend,
      surface: this.snapshot.surface,
      webGpuReady: backend !== null,
    })
  }

  private destroyCurrentBackend(): void {
    this.initToken += 1
    const backend = this.backend
    this.backend = null
    if (backend) {
      this.destroyBackend(backend)
    }
    this.updateSnapshot({
      backend: null,
      surface: this.snapshot.surface,
      webGpuReady: false,
    })
  }

  private syncCurrentSurface(): void {
    if (!this.active || !this.backend || !this.canvas || this.snapshot.surface.width <= 0 || this.snapshot.surface.height <= 0) {
      return
    }
    this.syncSurface({
      backend: this.backend,
      canvas: this.canvas,
      size: this.snapshot.surface,
    })
  }

  private updateSnapshot(snapshot: WorkbookPaneSurfaceSnapshotV3): void {
    this.snapshot = snapshot
    this.syncCurrentSurface()
    this.listener?.(snapshot)
  }
}
