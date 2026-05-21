export interface BiligRenderedCanvasState {
  readonly headerPaneCount: number
  readonly mode: string | null
  readonly pixelHeight: number
  readonly pixelWidth: number
  readonly tilePaneCount: number
  readonly visiblePixelCount?: number | undefined
}

export interface BiligRenderedSurfaceState {
  readonly dpr: number
  readonly fallback: BiligRenderedCanvasState | null
  readonly gridHeight: number
  readonly gridWidth: number
  readonly typeGpu: BiligRenderedCanvasState | null
}

export function isBiligRenderedSurfaceReady(state: BiligRenderedSurfaceState | null): boolean {
  if (!state || state.gridWidth <= 0 || state.gridHeight <= 0) {
    return false
  }
  if (state.fallback) {
    return false
  }
  const hasPaneData = (state.typeGpu?.tilePaneCount ?? 0) > 0 && (state.typeGpu?.headerPaneCount ?? 0) > 0
  if (!hasPaneData) {
    return false
  }
  return state.typeGpu?.mode === 'typegpu-v3' && canvasPixelsMatchViewport(state.typeGpu, state)
}

function canvasPixelsMatchViewport(canvas: BiligRenderedCanvasState, state: BiligRenderedSurfaceState): boolean {
  const expectedPixelWidth = Math.max(1, Math.floor(state.gridWidth * state.dpr))
  const expectedPixelHeight = Math.max(1, Math.floor(state.gridHeight * state.dpr))
  return canvas.pixelWidth >= expectedPixelWidth - 2 && canvas.pixelHeight >= expectedPixelHeight - 2
}
