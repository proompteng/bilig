import { VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'
import type { GridGpuScene } from '../gridGpuScene.js'
import type { GridTextScene } from '../gridTextScene.js'

export const GRID_SCENE_PACKET_V2_MAGIC = 'bilig.grid.scene.v2'
export const GRID_SCENE_PACKET_V2_VERSION = 1
export const GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT = 8
export const GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT = 20
export const GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT = 8
export type GridScenePacketPaneId = 'body' | 'top' | 'left' | 'corner' | 'top-frozen' | 'top-body' | 'left-frozen' | 'left-body' | 'overlay'
export type GridTilePaneKind =
  | 'body'
  | 'frozenTop'
  | 'frozenLeft'
  | 'frozenCorner'
  | 'columnHeaderBody'
  | 'columnHeaderFrozen'
  | 'rowHeaderBody'
  | 'rowHeaderFrozen'
  | 'dynamicOverlay'

export interface GridTileKeyV2 {
  readonly sheetName: string
  readonly paneKind: GridTilePaneKind
  readonly rowStart: number
  readonly rowEnd: number
  readonly colStart: number
  readonly colEnd: number
  readonly rowTile: number
  readonly colTile: number
  readonly axisVersionX: number
  readonly axisVersionY: number
  readonly valueVersion: number
  readonly styleVersion: number
  readonly selectionIndependentVersion: number
  readonly freezeVersion: number
  readonly textEpoch: number
  readonly dprBucket: number
}

export interface GridScenePacketV2 {
  readonly magic: typeof GRID_SCENE_PACKET_V2_MAGIC
  readonly version: typeof GRID_SCENE_PACKET_V2_VERSION
  readonly generation: number
  readonly key: GridTileKeyV2
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly rects: Float32Array
  readonly rectInstances: Float32Array
  readonly rectCount: number
  readonly fillRectCount: number
  readonly borderRectCount: number
  readonly textMetrics: Float32Array
  readonly textCount: number
}

export function packGridScenePacketV2(input: {
  readonly generation: number
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly key?: GridTileKeyV2 | undefined
  readonly surfaceSize: { readonly width: number; readonly height: number }
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}): GridScenePacketV2 {
  const key =
    input.key ??
    createGridTileKeyV2({
      paneId: input.paneId,
      sheetName: input.sheetName,
      viewport: input.viewport,
    })
  return {
    generation: input.generation,
    borderRectCount: input.gpuScene.borderRects.length,
    fillRectCount: input.gpuScene.fillRects.length,
    key,
    magic: GRID_SCENE_PACKET_V2_MAGIC,
    paneId: input.paneId,
    rectCount: input.gpuScene.fillRects.length + input.gpuScene.borderRects.length,
    rectInstances: packRectInstances(input.gpuScene, input.surfaceSize),
    rects: packRects(input.gpuScene),
    sheetName: input.sheetName,
    surfaceSize: input.surfaceSize,
    textCount: input.textScene.items.length,
    textMetrics: packTextMetrics(input.textScene),
    version: GRID_SCENE_PACKET_V2_VERSION,
    viewport: input.viewport,
  }
}

export function createGridTileKeyV2(input: {
  readonly sheetName: string
  readonly paneId: GridScenePacketPaneId
  readonly viewport: Viewport
  readonly rowTile?: number | undefined
  readonly colTile?: number | undefined
  readonly axisVersionX?: number | undefined
  readonly axisVersionY?: number | undefined
  readonly valueVersion?: number | undefined
  readonly styleVersion?: number | undefined
  readonly selectionIndependentVersion?: number | undefined
  readonly freezeVersion?: number | undefined
  readonly textEpoch?: number | undefined
  readonly dprBucket?: number | undefined
}): GridTileKeyV2 {
  return {
    axisVersionX: input.axisVersionX ?? 0,
    axisVersionY: input.axisVersionY ?? 0,
    colEnd: input.viewport.colEnd,
    colStart: input.viewport.colStart,
    colTile: input.colTile ?? Math.floor(input.viewport.colStart / VIEWPORT_TILE_COLUMN_COUNT),
    dprBucket: input.dprBucket ?? 1,
    freezeVersion: input.freezeVersion ?? 0,
    paneKind: resolveGridTilePaneKind(input.paneId),
    rowEnd: input.viewport.rowEnd,
    rowStart: input.viewport.rowStart,
    rowTile: input.rowTile ?? Math.floor(input.viewport.rowStart / VIEWPORT_TILE_ROW_COUNT),
    selectionIndependentVersion: input.selectionIndependentVersion ?? 0,
    sheetName: input.sheetName,
    styleVersion: input.styleVersion ?? 0,
    textEpoch: input.textEpoch ?? 0,
    valueVersion: input.valueVersion ?? 0,
  }
}

export function resolveGridTilePaneKind(paneId: GridScenePacketPaneId): GridTilePaneKind {
  switch (paneId) {
    case 'body':
      return 'body'
    case 'top':
      return 'frozenTop'
    case 'left':
      return 'frozenLeft'
    case 'corner':
      return 'frozenCorner'
    case 'top-body':
      return 'columnHeaderBody'
    case 'top-frozen':
      return 'columnHeaderFrozen'
    case 'left-body':
      return 'rowHeaderBody'
    case 'left-frozen':
      return 'rowHeaderFrozen'
    case 'overlay':
      return 'dynamicOverlay'
  }
}

export function serializeGridTileKeyV2(key: GridTileKeyV2): string {
  return [
    key.sheetName,
    key.paneKind,
    key.rowStart,
    key.rowEnd,
    key.colStart,
    key.colEnd,
    key.rowTile,
    key.colTile,
    key.axisVersionX,
    key.axisVersionY,
    key.valueVersion,
    key.styleVersion,
    key.selectionIndependentVersion,
    key.freezeVersion,
    key.textEpoch,
    key.dprBucket,
  ].join(':')
}

export function gridTileKeyV2Overlaps(left: GridTileKeyV2, right: GridTileKeyV2): boolean {
  return !(left.rowEnd < right.rowStart || right.rowEnd < left.rowStart || left.colEnd < right.colStart || right.colEnd < left.colStart)
}

export function isStaleValidGridTileKeyV2(candidate: GridTileKeyV2, desired: GridTileKeyV2): boolean {
  return (
    candidate.sheetName === desired.sheetName &&
    candidate.paneKind === desired.paneKind &&
    candidate.axisVersionX === desired.axisVersionX &&
    candidate.axisVersionY === desired.axisVersionY &&
    candidate.freezeVersion === desired.freezeVersion &&
    candidate.textEpoch === desired.textEpoch &&
    candidate.dprBucket === desired.dprBucket &&
    candidate.valueVersion <= desired.valueVersion &&
    candidate.styleVersion <= desired.styleVersion &&
    candidate.selectionIndependentVersion <= desired.selectionIndependentVersion &&
    gridTileKeyV2Overlaps(candidate, desired)
  )
}

function packRectInstances(scene: GridGpuScene, surfaceSize: { readonly width: number; readonly height: number }): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT)
  const clipX = 0
  const clipY = 0
  const clipX1 = surfaceSize.width
  const clipY1 = surfaceSize.height
  let offset = 0
  for (const rect of scene.fillRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    floats[offset + 8] = 0
    floats[offset + 9] = 0
    floats[offset + 10] = 0
    floats[offset + 11] = 0
    floats[offset + 12] = rect.color.a < 0.2 ? 2 : 0
    floats[offset + 13] = 0
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT
  }
  for (const rect of scene.borderRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = 0
    floats[offset + 5] = 0
    floats[offset + 6] = 0
    floats[offset + 7] = 0
    floats[offset + 8] = rect.color.r
    floats[offset + 9] = rect.color.g
    floats[offset + 10] = rect.color.b
    floats[offset + 11] = rect.color.a
    floats[offset + 12] = 0
    floats[offset + 13] = 1
    floats[offset + 14] = 0
    floats[offset + 15] = 0
    floats[offset + 16] = clipX
    floats[offset + 17] = clipY
    floats[offset + 18] = clipX1
    floats[offset + 19] = clipY1
    offset += GRID_SCENE_PACKET_V2_RECT_INSTANCE_FLOAT_COUNT
  }
  return floats
}

function packRects(scene: GridGpuScene): Float32Array {
  const rectCount = scene.fillRects.length + scene.borderRects.length
  const floats = new Float32Array(Math.max(1, rectCount) * GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT)
  let offset = 0
  for (const rect of scene.fillRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    offset += GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
  }
  for (const rect of scene.borderRects) {
    floats[offset + 0] = rect.x
    floats[offset + 1] = rect.y
    floats[offset + 2] = rect.width
    floats[offset + 3] = rect.height
    floats[offset + 4] = rect.color.r
    floats[offset + 5] = rect.color.g
    floats[offset + 6] = rect.color.b
    floats[offset + 7] = rect.color.a
    offset += GRID_SCENE_PACKET_V2_RECT_FLOAT_COUNT
  }
  return floats
}

function packTextMetrics(scene: GridTextScene): Float32Array {
  const floats = new Float32Array(Math.max(1, scene.items.length) * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT)
  scene.items.forEach((item, index) => {
    const offset = index * GRID_SCENE_PACKET_V2_TEXT_METRIC_FLOAT_COUNT
    floats[offset + 0] = item.x
    floats[offset + 1] = item.y
    floats[offset + 2] = item.width
    floats[offset + 3] = item.height
    floats[offset + 4] = item.clipInsetTop
    floats[offset + 5] = item.clipInsetRight
    floats[offset + 6] = item.clipInsetBottom
    floats[offset + 7] = item.clipInsetLeft
  })
  return floats
}
