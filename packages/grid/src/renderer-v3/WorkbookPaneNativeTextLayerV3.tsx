import { memo, useMemo, useSyncExternalStore, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'
import type { TextQuadRun } from './line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveTypeGpuV3DrawScrollSnapshot } from './workbook-pane-renderer-runtime.js'

type TextLayerPane = GridHeaderPaneState | WorkbookRenderTilePaneState

export interface WorkbookPaneNativeTextLayerV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
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

function isPaneDrawVisible(pane: TextLayerPane): boolean {
  return pane.drawVisible !== false
}

function getPaneTextRuns(pane: TextLayerPane): readonly TextQuadRun[] {
  return 'tile' in pane ? pane.tile.textRuns : pane.textRuns
}

function resolveNativeTextRunKey(pane: TextLayerPane, run: TextQuadRun): string {
  return [
    pane.paneId,
    run.text,
    run.x,
    run.y,
    run.width ?? '',
    run.height ?? '',
    run.clipX ?? '',
    run.clipY ?? '',
    run.clipWidth ?? '',
    run.clipHeight ?? '',
    run.font ?? '',
    run.color ?? '',
  ].join(':')
}

function getDefaultFont(run: TextQuadRun): string {
  return run.font?.trim()
    ? run.font
    : `400 ${run.fontSize ?? workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)}px ${WORKBOOK_FONT_SANS}`
}

function getDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1)
}

function snapCssPixel(value: number, dpr: number): number {
  return Math.round(value * dpr) / dpr
}

export function resolveNativeTextRunOuterStyleV3(input: {
  readonly pane: TextLayerPane
  readonly run: TextQuadRun
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly dpr?: number | undefined
}): CSSProperties {
  const dpr = input.dpr ?? getDevicePixelRatio()
  const offset = resolvePaneRenderOffset(input.pane, input.scrollSnapshot)
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clipX = input.run.clipX ?? input.run.x
  const clipY = input.run.clipY ?? input.run.y
  const clipWidth = input.run.clipWidth ?? width
  const clipHeight = input.run.clipHeight ?? height
  return {
    height: clipHeight,
    left: snapCssPixel(input.pane.frame.x + offset.x + clipX, dpr),
    overflow: 'hidden',
    position: 'absolute',
    top: snapCssPixel(input.pane.frame.y + offset.y + clipY, dpr),
    width: clipWidth,
  }
}

export function resolveNativeTextRunInnerStyleV3(input: { readonly run: TextQuadRun; readonly dpr?: number | undefined }): CSSProperties {
  const dpr = input.dpr ?? getDevicePixelRatio()
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clipX = input.run.clipX ?? input.run.x
  const clipY = input.run.clipY ?? input.run.y
  const justifyContent = input.run.align === 'right' ? 'flex-end' : input.run.align === 'center' ? 'center' : 'flex-start'
  return {
    alignItems: input.run.wrap ? 'flex-start' : 'center',
    boxSizing: 'border-box',
    color: input.run.color ?? '#111827',
    display: 'flex',
    font: getDefaultFont(input.run),
    fontKerning: 'normal',
    height,
    justifyContent,
    left: snapCssPixel(input.run.x - clipX, dpr),
    lineHeight: 1.2,
    overflow: 'hidden',
    paddingLeft: 6,
    paddingRight: 6,
    position: 'absolute',
    textAlign: input.run.align ?? 'left',
    textDecorationLine: input.run.underline ? 'underline' : input.run.strike ? 'line-through' : undefined,
    textRendering: 'auto',
    top: snapCssPixel(input.run.y - clipY, dpr),
    whiteSpace: input.run.wrap ? 'pre-wrap' : 'pre',
    width,
    WebkitFontSmoothing: 'auto',
  }
}

export const WorkbookPaneNativeTextLayerV3 = memo(function WorkbookPaneNativeTextLayerV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes,
  scrollTransformStore,
  tilePanes,
}: WorkbookPaneNativeTextLayerV3Props) {
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
      resolveTypeGpuV3DrawScrollSnapshot({
        fallback: scrollSnapshot,
        geometry: resolvedGeometry,
        panes: tilePanes,
      }),
    [resolvedGeometry, scrollSnapshot, tilePanes],
  )
  const panes = useMemo<readonly TextLayerPane[]>(() => [...tilePanes, ...headerPanes], [headerPanes, tilePanes])
  const dpr = getDevicePixelRatio()
  const textRunCount = panes.reduce((total, pane) => total + getPaneTextRuns(pane).length, 0)

  if (!active || textRunCount === 0) {
    return null
  }

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-[15] overflow-hidden"
      data-v3-native-text-layer="mounted"
      data-v3-native-text-run-count={textRunCount}
      data-testid="grid-native-text-layer"
      style={{ contain: 'strict' }}
    >
      {panes.flatMap((pane) => {
        if (!isPaneDrawVisible(pane)) {
          return []
        }
        return getPaneTextRuns(pane).flatMap((run) => {
          const width = run.width ?? 0
          const height = run.height ?? 0
          const clipWidth = run.clipWidth ?? width
          const clipHeight = run.clipHeight ?? height
          if (!run.text || clipWidth <= 0 || clipHeight <= 0) {
            return []
          }
          return (
            <div
              data-native-text-run=""
              key={resolveNativeTextRunKey(pane, run)}
              style={resolveNativeTextRunOuterStyleV3({ dpr, pane, run, scrollSnapshot: drawScrollSnapshot })}
            >
              <div style={resolveNativeTextRunInnerStyleV3({ dpr, run })}>{run.text}</div>
            </div>
          )
        })
      })}
    </div>
  )
})
