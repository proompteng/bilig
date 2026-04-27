import { GridRenderLoop } from './gridRenderLoop.js'

export const TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS = 48
export const TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS = 64

export interface GridDrawSchedulerCameraV3 {
  readonly updatedAt: number
  readonly velocityX: number
  readonly velocityY: number
}

export interface GridDrawSchedulerFrameDecisionV3 {
  readonly deferPreloadSync: boolean
  readonly syncPreloadPanes: boolean
}

export function shouldDeferTypeGpuV3PreloadSync(input: {
  readonly now: number
  readonly lastScrollSignalAt: number
  readonly camera: GridDrawSchedulerCameraV3 | null
}): boolean {
  const hasMovingCamera =
    input.camera !== null &&
    Math.abs(input.camera.velocityX) + Math.abs(input.camera.velocityY) > 0.01 &&
    input.now - input.camera.updatedAt < TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS
  return hasMovingCamera || input.now - input.lastScrollSignalAt < TYPEGPU_V3_ACTIVE_RESOURCE_DEFER_MS
}

export class GridDrawSchedulerV3 {
  private readonly renderLoop: GridRenderLoop
  private idlePreloadRetry: number | null = null
  private lastScrollSignalAt = 0

  constructor(
    private readonly scheduleIdle: (callback: () => void, delay: number) => number = window.setTimeout.bind(window),
    private readonly cancelIdle: (handle: number) => void = window.clearTimeout.bind(window),
    private readonly now: () => number = () => performance.now(),
    renderLoop = new GridRenderLoop(),
  ) {
    this.renderLoop = renderLoop
  }

  noteInputSignal(timestamp = this.now()): void {
    this.lastScrollSignalAt = timestamp
  }

  requestDraw(draw: () => void): void {
    this.renderLoop.requestDraw(draw)
  }

  resolveFrame(input: {
    readonly camera: GridDrawSchedulerCameraV3 | null
    readonly requestIdlePreloadDraw: () => void
  }): GridDrawSchedulerFrameDecisionV3 {
    const deferPreloadSync = shouldDeferTypeGpuV3PreloadSync({
      camera: input.camera,
      lastScrollSignalAt: this.lastScrollSignalAt,
      now: this.now(),
    })
    if (deferPreloadSync) {
      this.scheduleIdlePreloadRetry(input.requestIdlePreloadDraw)
    }
    return {
      deferPreloadSync,
      syncPreloadPanes: !deferPreloadSync,
    }
  }

  cancel(): void {
    this.renderLoop.cancel()
    if (this.idlePreloadRetry !== null) {
      this.cancelIdle(this.idlePreloadRetry)
      this.idlePreloadRetry = null
    }
  }

  private scheduleIdlePreloadRetry(requestIdlePreloadDraw: () => void): void {
    if (this.idlePreloadRetry !== null) {
      this.cancelIdle(this.idlePreloadRetry)
    }
    this.idlePreloadRetry = this.scheduleIdle(() => {
      this.idlePreloadRetry = null
      requestIdlePreloadDraw()
    }, TYPEGPU_V3_IDLE_PRELOAD_RETRY_MS)
  }
}
