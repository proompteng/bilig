import { memo, useMemo, useSyncExternalStore, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from './rect-instance-buffer.js'
import { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'

type RectLayerPane = GridHeaderPaneState | WorkbookRenderTilePaneState

export interface NativeRectLayerRectV3 {
  readonly alpha: number
  readonly color: string
  readonly height: number
  readonly key: string
  readonly left: number
  readonly top: number
  readonly width: number
}

export interface WorkbookPaneNativeRectLayerV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly presentedScrollSnapshot?: WorkbookGridScrollSnapshot | null | undefined
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

const EMPTY_SCROLL_SNAPSHOT: WorkbookGridScrollSnapshot = Object.freeze({ tx: 0, ty: 0 })

function subscribeNoop(): () => void {
  return () => {}
}

function getNullSnapshot(): GridGeometrySnapshot | null {
  return null
}

function isPaneDrawVisible(pane: RectLayerPane): boolean {
  return pane.drawVisible !== false
}

function getPaneRectInstances(pane: RectLayerPane): Float32Array {
  return 'tile' in pane ? pane.tile.rectInstances : pane.rectInstances
}

function getPaneRectCount(pane: RectLayerPane): number {
  return 'tile' in pane ? pane.tile.rectCount : pane.rectCount
}

function resolvePaneRenderOffset(
  pane: {
    readonly contentOffset: { readonly x: number; readonly y: number }
    readonly scrollAxes: { readonly x: boolean; readonly y: boolean }
  },
  scrollSnapshot: {
    readonly tx: number
    readonly ty: number
    readonly renderTx?: number | undefined
    readonly renderTy?: number | undefined
  },
): { readonly x: number; readonly y: number } {
  const renderTx = scrollSnapshot.renderTx ?? scrollSnapshot.tx
  const renderTy = scrollSnapshot.renderTy ?? scrollSnapshot.ty
  return {
    x: pane.contentOffset.x - (pane.scrollAxes.x ? renderTx : 0),
    y: pane.contentOffset.y - (pane.scrollAxes.y ? renderTy : 0),
  }
}

function resolveLatestGeometry(
  propGeometry: GridGeometrySnapshot | null,
  liveGeometry: GridGeometrySnapshot | null,
): GridGeometrySnapshot | null {
  if (!propGeometry) {
    return liveGeometry
  }
  if (!liveGeometry) {
    return propGeometry
  }
  return liveGeometry.camera.seq > propGeometry.camera.seq ? liveGeometry : propGeometry
}

function resolveNativeRectLayerDrawScrollSnapshotV3(input: {
  readonly geometry: GridGeometrySnapshot | null
  readonly liveScrollSnapshot: WorkbookGridScrollSnapshot
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly presentedScrollSnapshot?: WorkbookGridScrollSnapshot | null | undefined
}): WorkbookGridScrollSnapshot {
  if (input.presentedScrollSnapshot) {
    return input.presentedScrollSnapshot
  }
  return resolveTypeGpuV3DrawScrollSnapshot({
    fallback: input.liveScrollSnapshot,
    geometry: input.geometry,
    panes: input.panes,
  })
}

export function buildNativeRectLayerRectsForPaneV3(input: {
  readonly pane: RectLayerPane
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
}): readonly NativeRectLayerRectV3[] {
  if (!isPaneDrawVisible(input.pane)) {
    return []
  }
  const rectCount = getPaneRectCount(input.pane)
  if (rectCount <= 0) {
    return []
  }
  const instances = getPaneRectInstances(input.pane)
  const offset = resolvePaneRenderOffset(input.pane, input.scrollSnapshot)
  const rects: NativeRectLayerRectV3[] = []
  const paneKey = resolvePaneIdentity(input.pane)
  const maxOffset = Math.min(rectCount * GRID_RECT_INSTANCE_FLOAT_COUNT_V3, instances.length)
  for (let rectOffset = 0; rectOffset + GRID_RECT_INSTANCE_FLOAT_COUNT_V3 <= maxOffset; rectOffset += GRID_RECT_INSTANCE_FLOAT_COUNT_V3) {
    const width = instances[rectOffset + 2] ?? 0
    const height = instances[rectOffset + 3] ?? 0
    if (width <= 0 || height <= 0) {
      continue
    }
    const borderAlpha = instances[rectOffset + 11] ?? 0
    const borderThickness = instances[rectOffset + 13] ?? 0
    const colorOffset = borderThickness > 0 && borderAlpha > 0 ? rectOffset + 8 : rectOffset + 4
    const alpha = instances[colorOffset + 3] ?? 0
    if (alpha <= 0.001) {
      continue
    }
    const unclippedLeft = offset.x + (instances[rectOffset] ?? 0)
    const unclippedTop = offset.y + (instances[rectOffset + 1] ?? 0)
    const right = unclippedLeft + width
    const bottom = unclippedTop + height
    if (right <= 0 || bottom <= 0 || unclippedLeft >= input.pane.frame.width || unclippedTop >= input.pane.frame.height) {
      continue
    }
    const left = Math.max(0, unclippedLeft)
    const top = Math.max(0, unclippedTop)
    const clippedWidth = Math.min(input.pane.frame.width, right) - left
    const clippedHeight = Math.min(input.pane.frame.height, bottom) - top
    if (clippedWidth <= 0 || clippedHeight <= 0) {
      continue
    }
    rects.push({
      alpha,
      color: rgbaCssColor(instances[colorOffset] ?? 0, instances[colorOffset + 1] ?? 0, instances[colorOffset + 2] ?? 0, alpha),
      height: clippedHeight,
      key: `${paneKey}:${rectOffset}:${left}:${top}:${clippedWidth}:${clippedHeight}:${colorOffset}:${alpha}`,
      left,
      top,
      width: clippedWidth,
    })
  }
  return rects
}

export const WorkbookPaneNativeRectLayerV3 = memo(function WorkbookPaneNativeRectLayerV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes,
  presentedScrollSnapshot = null,
  scrollTransformStore,
  tilePanes,
}: WorkbookPaneNativeRectLayerV3Props) {
  const liveGeometry = useSyncExternalStore(
    cameraStore ? cameraStore.subscribe.bind(cameraStore) : subscribeNoop,
    cameraStore ? cameraStore.getSnapshot.bind(cameraStore) : getNullSnapshot,
    getNullSnapshot,
  )
  const scrollSnapshot = useSyncExternalStore(
    scrollTransformStore ? scrollTransformStore.subscribe.bind(scrollTransformStore) : subscribeNoop,
    scrollTransformStore ? scrollTransformStore.getSnapshot.bind(scrollTransformStore) : () => EMPTY_SCROLL_SNAPSHOT,
    () => EMPTY_SCROLL_SNAPSHOT,
  )
  const resolvedGeometry = resolveLatestGeometry(geometry, liveGeometry)
  const drawScrollSnapshot = useMemo(
    () =>
      resolveNativeRectLayerDrawScrollSnapshotV3({
        geometry: resolvedGeometry,
        liveScrollSnapshot: scrollSnapshot,
        panes: tilePanes,
        presentedScrollSnapshot,
      }),
    [presentedScrollSnapshot, resolvedGeometry, scrollSnapshot, tilePanes],
  )
  const panes = useMemo<readonly RectLayerPane[]>(() => [...tilePanes, ...headerPanes], [headerPanes, tilePanes])
  const renderedPanes = useMemo(
    () =>
      active
        ? panes
            .map((pane) => ({
              pane,
              rects: buildNativeRectLayerRectsForPaneV3({ pane, scrollSnapshot: drawScrollSnapshot }),
            }))
            .filter((entry) => entry.rects.length > 0)
        : [],
    [active, drawScrollSnapshot, panes],
  )
  const rectCount = renderedPanes.reduce((total, entry) => total + entry.rects.length, 0)
  if (!active || rectCount === 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-[12] overflow-hidden"
      data-testid="grid-native-rect-layer"
      data-v3-native-rect-layer="mounted"
      data-v3-native-rect-count={rectCount}
      style={{ contain: 'strict' }}
    >
      {renderedPanes.map(({ pane, rects }) => (
        <div data-native-rect-pane={resolvePaneIdentity(pane)} key={resolvePaneIdentity(pane)} style={resolvePaneStyle(pane.frame)}>
          {rects.map((rect) => (
            <div data-native-rect="" key={rect.key} style={resolveRectStyle(rect)} />
          ))}
        </div>
      ))}
    </div>
  )
})

function resolvePaneIdentity(pane: RectLayerPane): string {
  return 'tile' in pane
    ? [pane.paneId, pane.tile.tileId, pane.tile.coord.sheetOrdinal, pane.tile.coord.rowTile, pane.tile.coord.colTile].join(':')
    : pane.paneId
}

function resolvePaneStyle(frame: {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}): CSSProperties {
  return {
    contain: 'strict',
    height: frame.height,
    left: frame.x,
    overflow: 'hidden',
    position: 'absolute',
    top: frame.y,
    width: frame.width,
  }
}

function resolveRectStyle(rect: NativeRectLayerRectV3): CSSProperties {
  return {
    backgroundColor: rect.color,
    height: rect.height,
    left: rect.left,
    position: 'absolute',
    top: rect.top,
    width: rect.width,
  }
}

function rgbaCssColor(r: number, g: number, b: number, a: number): string {
  const red = Math.round(clamp01(r) * 255)
  const green = Math.round(clamp01(g) * 255)
  const blue = Math.round(clamp01(b) * 255)
  return `rgba(${red}, ${green}, ${blue}, ${clamp01(a).toFixed(4)})`
}

function clamp01(value: number): number {
  if (!Number.isFinite(value)) {
    return 0
  }
  return Math.max(0, Math.min(1, value))
}
