import type { WorkbookGridScrollSnapshot } from '../workbookGridScrollStore.js'
import type { WorkbookRenderPaneState } from '../renderer/pane-scene-types.js'
import { createGlyphAtlas } from './typegpu-atlas-manager.js'
import { WorkbookPaneBufferCache } from './pane-buffer-cache.js'
import {
  createTypeGpuRenderer,
  destroyTypeGpuRenderer,
  syncTypeGpuAtlasResources,
  type TypeGpuRendererArtifacts,
} from './typegpu-backend.js'
import { drawTypeGpuPanes, type TypeGpuDrawSurface } from './typegpu-render-pass.js'
import { syncTypeGpuPaneResources } from './typegpu-buffer-pool.js'
import { createTypeGpuSurfaceState, syncTypeGpuCanvasSurface, type TypeGpuSurfaceState } from './typegpu-surface.js'
import { TileGpuCache, syncTileGpuCacheFromPanes } from './tile-gpu-cache.js'

export interface WorkbookTypeGpuBackend {
  readonly artifacts: TypeGpuRendererArtifacts
  readonly atlas: ReturnType<typeof createGlyphAtlas>
  readonly paneBuffers: WorkbookPaneBufferCache
  readonly surfaceState: TypeGpuSurfaceState
  readonly tileCache: TileGpuCache
}

export async function createWorkbookTypeGpuBackend(canvas: HTMLCanvasElement): Promise<WorkbookTypeGpuBackend | null> {
  const artifacts = await createTypeGpuRenderer(canvas)
  if (!artifacts) {
    return null
  }
  return {
    artifacts,
    atlas: createGlyphAtlas(),
    paneBuffers: new WorkbookPaneBufferCache(),
    surfaceState: createTypeGpuSurfaceState(),
    tileCache: new TileGpuCache(),
  }
}

export function destroyWorkbookTypeGpuBackend(backend: WorkbookTypeGpuBackend): void {
  backend.paneBuffers.dispose()
  destroyTypeGpuRenderer(backend.artifacts)
}

export function syncWorkbookTypeGpuSurface(input: {
  readonly backend: WorkbookTypeGpuBackend
  readonly canvas: HTMLCanvasElement
  readonly size: TypeGpuDrawSurface
}): void {
  syncTypeGpuCanvasSurface({
    artifacts: input.backend.artifacts,
    canvas: input.canvas,
    size: input.size,
    state: input.backend.surfaceState,
  })
}

export function drawWorkbookTypeGpuFrame(input: {
  readonly backend: WorkbookTypeGpuBackend
  readonly panes: readonly WorkbookRenderPaneState[]
  readonly deferTextUploads?: boolean | undefined
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly surface: TypeGpuDrawSurface
}): void {
  syncTileGpuCacheFromPanes({
    cache: input.backend.tileCache,
    panes: input.panes,
  })
  syncTypeGpuPaneResources({
    artifacts: input.backend.artifacts,
    atlas: input.backend.atlas,
    paneBuffers: input.backend.paneBuffers,
    panes: input.panes,
    deferTextUploads: input.deferTextUploads,
  })
  syncTypeGpuAtlasResources(input.backend.artifacts, input.backend.atlas)
  drawTypeGpuPanes({
    artifacts: input.backend.artifacts,
    paneBuffers: input.backend.paneBuffers,
    panes: input.panes,
    scrollSnapshot: input.scrollSnapshot,
    surface: input.surface,
  })
}
