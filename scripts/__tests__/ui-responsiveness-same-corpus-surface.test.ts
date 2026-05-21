import { describe, expect, it } from 'vitest'

import { isBiligRenderedSurfaceReady, type BiligRenderedSurfaceState } from '../ui-responsiveness-same-corpus-surface.ts'

const readyTypeGpuSurface: BiligRenderedSurfaceState = {
  dpr: 2,
  fallback: null,
  gridHeight: 300,
  gridWidth: 500,
  typeGpu: {
    headerPaneCount: 1,
    mode: 'typegpu-v3',
    pixelHeight: 600,
    pixelWidth: 1000,
    tilePaneCount: 1,
  },
}

describe('same-corpus Bilig rendered surface proof', () => {
  it('requires the TypeGPU renderer without accepting a fallback canvas', () => {
    expect(isBiligRenderedSurfaceReady(readyTypeGpuSurface)).toBe(true)
    expect(
      isBiligRenderedSurfaceReady({
        ...readyTypeGpuSurface,
        fallback: {
          headerPaneCount: 1,
          mode: 'legacy-fallback',
          pixelHeight: 600,
          pixelWidth: 1000,
          tilePaneCount: 1,
        },
      }),
    ).toBe(false)
  })
})
