import {
  createWorkbookTypeGpuBackendV3,
  destroyWorkbookTypeGpuBackendV3,
  syncWorkbookTypeGpuSurfaceV3,
  type WorkbookTypeGpuBackendV3,
} from './typegpu-workbook-backend-v3.js'
import type { TypeGpuSurfaceSizeV3 } from './workbook-pane-renderer-runtime.js'

export interface WorkbookPaneSurfaceSnapshotV3 {
  readonly backend: object | null
  readonly backendStatus: WorkbookPaneSurfaceBackendStatusV3
  readonly surface: TypeGpuSurfaceSizeV3
  readonly webGpuReady: boolean
}

export type WorkbookPaneSurfaceBackendStatusV3 = 'idle' | 'initializing' | 'ready' | 'unavailable'

export interface WorkbookPaneSurfaceRuntimeResizeObserverV3 {
  observe(target: Element): void
  disconnect(): void
}

export interface WorkbookPaneSurfaceResizeEntryV3 {
  readonly target?: Element | undefined
  readonly devicePixelContentBoxSize?: readonly ResizeObserverSize[] | ResizeObserverSize | undefined
}

export type WorkbookPaneSurfaceResizeListenerV3 = (entries: readonly WorkbookPaneSurfaceResizeEntryV3[]) => void

export interface WorkbookPaneSurfaceRuntimeOptionsV3 {
  readonly createBackend?: ((canvas: HTMLCanvasElement) => Promise<object | null>) | undefined
  readonly destroyBackend?: ((backend: object) => void) | undefined
  readonly syncSurface?:
    | ((input: { readonly backend: object; readonly canvas: HTMLCanvasElement; readonly size: TypeGpuSurfaceSizeV3 }) => void)
    | undefined
  readonly getDevicePixelRatio?: (() => number) | undefined
  readonly createResizeObserver?:
    | ((listener: WorkbookPaneSurfaceResizeListenerV3) => WorkbookPaneSurfaceRuntimeResizeObserverV3 | null)
    | undefined
}

const EMPTY_TYPEGPU_SURFACE_SIZE_V3: TypeGpuSurfaceSizeV3 = Object.freeze({
  dpr: 1,
  height: 0,
  pixelHeight: 0,
  pixelWidth: 0,
  width: 0,
})

export const EMPTY_WORKBOOK_PANE_SURFACE_SNAPSHOT_V3: WorkbookPaneSurfaceSnapshotV3 = Object.freeze({
  backend: null,
  backendStatus: 'idle',
  surface: EMPTY_TYPEGPU_SURFACE_SIZE_V3,
  webGpuReady: false,
})

export function resolveWorkbookPaneSurfaceSizeV3(input: {
  readonly host: Pick<HTMLElement, 'clientHeight' | 'clientWidth'>
  readonly dpr?: number | undefined
  readonly resizeEntry?: WorkbookPaneSurfaceResizeEntryV3
}): TypeGpuSurfaceSizeV3 {
  const width = Math.max(0, Math.floor(input.host.clientWidth))
  const height = Math.max(0, Math.floor(input.host.clientHeight))
  const dpr = Math.max(1, input.dpr ?? defaultDevicePixelRatio())
  const devicePixelSize = resolveResizeObserverDevicePixelSizeV3(input.resizeEntry)
  const fallbackPixelWidth = Math.max(1, Math.ceil(width * dpr))
  const fallbackPixelHeight = Math.max(1, Math.ceil(height * dpr))
  return {
    dpr,
    height,
    pixelHeight: resolveSurfacePixelSizeV3(fallbackPixelHeight, devicePixelSize?.height, dpr),
    pixelWidth: resolveSurfacePixelSizeV3(fallbackPixelWidth, devicePixelSize?.width, dpr),
    width,
  }
}

function resolveSurfacePixelSizeV3(fallback: number, devicePixelSize: number | undefined, dpr: number): number {
  if (typeof devicePixelSize !== 'number' || !Number.isFinite(devicePixelSize)) {
    return fallback
  }
  const tolerance = Math.max(1, Math.ceil(dpr))
  return Math.abs(devicePixelSize - fallback) <= tolerance ? devicePixelSize : fallback
}

function resolveResizeObserverDevicePixelSizeV3(
  entry: WorkbookPaneSurfaceResizeEntryV3 | undefined,
): { readonly height: number; readonly width: number } | null {
  const boxSize = entry?.devicePixelContentBoxSize
  const firstSize = Array.isArray(boxSize) ? boxSize[0] : boxSize
  const width = firstSize?.inlineSize
  const height = firstSize?.blockSize
  if (!Number.isFinite(width) || !Number.isFinite(height)) {
    return null
  }
  const pixelWidth = Math.max(1, Math.round(width))
  const pixelHeight = Math.max(1, Math.round(height))
  return pixelWidth > 0 && pixelHeight > 0 ? { height: pixelHeight, width: pixelWidth } : null
}

function defaultDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1
}

function createDefaultResizeObserver(listener: WorkbookPaneSurfaceResizeListenerV3): WorkbookPaneSurfaceRuntimeResizeObserverV3 | null {
  if (typeof ResizeObserver === 'undefined') {
    return null
  }
  return new ResizeObserver((entries) => listener(entries))
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
  private backendStatus: WorkbookPaneSurfaceBackendStatusV3 = 'idle'
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
  private readonly createResizeObserver: (
    listener: WorkbookPaneSurfaceResizeListenerV3,
  ) => WorkbookPaneSurfaceRuntimeResizeObserverV3 | null

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
      this.destroyCurrentBackend('idle')
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
    const observer = this.createResizeObserver((entries) => {
      const resizeEntry = entries.find((entry) => entry.target === this.host) ?? entries[0]
      this.refreshSurface(resizeEntry)
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

  private refreshSurface(resizeEntry?: WorkbookPaneSurfaceResizeEntryV3): void {
    const surface = this.host
      ? resolveWorkbookPaneSurfaceSizeV3({
          dpr: this.getDevicePixelRatio(),
          host: this.host,
          ...(resizeEntry ? { resizeEntry } : {}),
        })
      : EMPTY_TYPEGPU_SURFACE_SIZE_V3
    this.updateSnapshot({
      backend: this.backend,
      backendStatus: this.backendStatus,
      surface,
      webGpuReady: this.backend !== null,
    })
  }

  private async ensureBackend(): Promise<void> {
    const canvas = this.canvas
    if (!this.active || !canvas || this.backend || this.backendStatus === 'initializing' || this.backendStatus === 'unavailable') {
      return
    }
    const token = ++this.initToken
    this.backendStatus = 'initializing'
    this.refreshSurface()
    let backend: object | null = null
    try {
      backend = await this.createBackend(canvas)
    } catch {
      backend = null
    }
    if (token !== this.initToken || !this.active || this.canvas !== canvas) {
      if (backend) {
        this.destroyBackend(backend)
      }
      return
    }
    this.backend = backend
    this.backendStatus = backend ? 'ready' : 'unavailable'
    this.updateSnapshot({
      backend,
      backendStatus: this.backendStatus,
      surface: this.snapshot.surface,
      webGpuReady: backend !== null,
    })
  }

  private destroyCurrentBackend(nextStatus: WorkbookPaneSurfaceBackendStatusV3 = 'idle'): void {
    this.initToken += 1
    const backend = this.backend
    this.backend = null
    this.backendStatus = nextStatus
    if (backend) {
      this.destroyBackend(backend)
    }
    this.updateSnapshot({
      backend: null,
      backendStatus: this.backendStatus,
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
