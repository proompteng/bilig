import type { TypeGpuRendererArtifacts } from './typegpu-backend.js'
import { noteTypeGpuConfigure, noteTypeGpuSurfaceResize } from './grid-render-counters.js'

export interface TypeGpuSurfaceSize {
  readonly width: number
  readonly height: number
  readonly pixelWidth: number
  readonly pixelHeight: number
  readonly dpr: number
}

export interface TypeGpuSurfaceState {
  width: number
  height: number
  pixelWidth: number
  pixelHeight: number
  dpr: number
}

export function createTypeGpuSurfaceState(): TypeGpuSurfaceState {
  return {
    dpr: 0,
    height: 0,
    pixelHeight: 0,
    pixelWidth: 0,
    width: 0,
  }
}

export function syncTypeGpuCanvasSurface(input: {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly canvas: HTMLCanvasElement
  readonly size: TypeGpuSurfaceSize
  readonly state: TypeGpuSurfaceState
}): void {
  const { artifacts, canvas, size, state } = input
  if (size.width <= 0 || size.height <= 0 || size.pixelWidth <= 0 || size.pixelHeight <= 0) {
    return
  }

  const changed =
    state.width !== size.width ||
    state.height !== size.height ||
    state.pixelWidth !== size.pixelWidth ||
    state.pixelHeight !== size.pixelHeight ||
    state.dpr !== size.dpr
  if (!changed) {
    return
  }

  if (canvas.width !== size.pixelWidth) {
    canvas.width = size.pixelWidth
  }
  if (canvas.height !== size.pixelHeight) {
    canvas.height = size.pixelHeight
  }
  const cssWidth = `${size.width}px`
  if (canvas.style.width !== cssWidth) {
    canvas.style.width = cssWidth
  }
  const cssHeight = `${size.height}px`
  if (canvas.style.height !== cssHeight) {
    canvas.style.height = cssHeight
  }

  artifacts.context.configure({
    alphaMode: 'premultiplied',
    device: artifacts.device,
    format: artifacts.format,
  })
  noteTypeGpuConfigure()
  noteTypeGpuSurfaceResize(size.width, size.height, size.dpr)

  state.dpr = size.dpr
  state.height = size.height
  state.pixelHeight = size.pixelHeight
  state.pixelWidth = size.pixelWidth
  state.width = size.width
}
