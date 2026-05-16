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

interface NativeTextRunVisibleClipV3 {
  readonly outerHeight: number
  readonly outerLeft: number
  readonly outerTop: number
  readonly outerWidth: number
  readonly innerHeight: number
  readonly innerLeft: number
  readonly innerTop: number
  readonly innerWidth: number
}

export interface NativeTextRunFontStyleV3 {
  readonly fontFamily: string
  readonly fontSize: number
  readonly fontStyle: 'italic' | 'normal'
  readonly fontWeight: number | 'normal' | 'bold'
}

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

export function resolveNativeTextRunFontStyleV3(run: TextQuadRun): NativeTextRunFontStyleV3 {
  const font = getDefaultFont(run)
  const sizeCssPx = Math.max(1, run.fontSize ?? parseNativeTextRunFontSize(font))
  const sizeMatch = font.match(/\b\d+(?:\.\d+)?px\s+(.+)$/i)
  const fontFamily = sizeMatch?.[1]?.trim() || WORKBOOK_FONT_SANS
  const fontStyle = /\bitalic\b/i.test(font) ? 'italic' : 'normal'
  const weightMatch = font.match(/\b([1-9]00|normal|bold)\b/i)
  const rawWeight = weightMatch?.[1]?.toLowerCase()
  const fontWeight = rawWeight === 'bold' ? 'bold' : rawWeight === 'normal' || rawWeight === undefined ? 400 : Number(rawWeight)
  return {
    fontFamily,
    fontSize: sizeCssPx,
    fontStyle,
    fontWeight,
  }
}

function parseNativeTextRunFontSize(font: string): number {
  const match = font.match(/\b(\d+(?:\.\d+)?)px\b/i)
  return match ? Number(match[1]) : workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
}

function getDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1)
}

function snapCssPixel(value: number, dpr: number): number {
  return Math.round(value * dpr) / dpr
}

export function resolveNativeTextRunVisibleClipV3(input: {
  readonly pane: TextLayerPane
  readonly run: TextQuadRun
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly dpr?: number | undefined
}): NativeTextRunVisibleClipV3 | null {
  const dpr = input.dpr ?? getDevicePixelRatio()
  const offset = resolvePaneRenderOffset(input.pane, input.scrollSnapshot)
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clipX = input.run.clipX ?? input.run.x
  const clipY = input.run.clipY ?? input.run.y
  const clipWidth = input.run.clipWidth ?? width
  const clipHeight = input.run.clipHeight ?? height
  const clipLeft = input.pane.frame.x + offset.x + clipX
  const clipTop = input.pane.frame.y + offset.y + clipY
  const clipRight = clipLeft + clipWidth
  const clipBottom = clipTop + clipHeight
  const paneLeft = input.pane.frame.x
  const paneTop = input.pane.frame.y
  const paneRight = paneLeft + input.pane.frame.width
  const paneBottom = paneTop + input.pane.frame.height
  const visibleLeft = Math.max(clipLeft, paneLeft)
  const visibleTop = Math.max(clipTop, paneTop)
  const visibleRight = Math.min(clipRight, paneRight)
  const visibleBottom = Math.min(clipBottom, paneBottom)
  if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) {
    return null
  }

  const contentLeft = input.pane.frame.x + offset.x + input.run.x
  const contentTop = input.pane.frame.y + offset.y + input.run.y
  const outerLeft = snapCssPixel(visibleLeft, dpr)
  const outerTop = snapCssPixel(visibleTop, dpr)
  const outerRight = snapCssPixel(visibleRight, dpr)
  const outerBottom = snapCssPixel(visibleBottom, dpr)
  const innerRight = Math.min(contentLeft + width, visibleRight)
  const innerBottom = Math.min(contentTop + height, visibleBottom)

  return {
    innerHeight: Math.max(0, snapCssPixel(innerBottom - contentTop, dpr)),
    innerLeft: snapCssPixel(contentLeft - visibleLeft, dpr),
    innerTop: snapCssPixel(contentTop - visibleTop, dpr),
    innerWidth: Math.max(0, snapCssPixel(innerRight - contentLeft, dpr)),
    outerHeight: Math.max(0, outerBottom - outerTop),
    outerLeft,
    outerTop,
    outerWidth: Math.max(0, outerRight - outerLeft),
  }
}

export function resolveNativeTextRunOuterStyleV3(input: {
  readonly pane: TextLayerPane
  readonly run: TextQuadRun
  readonly scrollSnapshot: WorkbookGridScrollSnapshot
  readonly dpr?: number | undefined
  readonly visibleClip?: NativeTextRunVisibleClipV3 | null | undefined
}): CSSProperties {
  const visibleClip = input.visibleClip ?? resolveNativeTextRunVisibleClipV3(input)
  if (!visibleClip) {
    return { display: 'none' }
  }
  return {
    height: visibleClip.outerHeight,
    left: visibleClip.outerLeft,
    overflow: 'hidden',
    position: 'absolute',
    top: visibleClip.outerTop,
    width: visibleClip.outerWidth,
  }
}

export function resolveNativeTextRunInnerStyleV3(input: {
  readonly run: TextQuadRun
  readonly dpr?: number | undefined
  readonly visibleClip?: NativeTextRunVisibleClipV3 | null | undefined
}): CSSProperties {
  const dpr = input.dpr ?? getDevicePixelRatio()
  const width = input.run.width ?? 0
  const height = input.run.height ?? 0
  const clipX = input.run.clipX ?? input.run.x
  const clipY = input.run.clipY ?? input.run.y
  const visibleClip = input.visibleClip ?? null
  const fontStyle = resolveNativeTextRunFontStyleV3(input.run)
  const justifyContent = input.run.align === 'right' ? 'flex-end' : input.run.align === 'center' ? 'center' : 'flex-start'
  return {
    alignItems: input.run.wrap ? 'flex-start' : 'center',
    boxSizing: 'border-box',
    color: input.run.color ?? '#111827',
    display: 'flex',
    fontFamily: fontStyle.fontFamily,
    fontSize: fontStyle.fontSize,
    fontStyle: fontStyle.fontStyle,
    fontKerning: 'normal',
    fontWeight: fontStyle.fontWeight,
    height: visibleClip?.innerHeight ?? height,
    justifyContent,
    left: visibleClip?.innerLeft ?? snapCssPixel(input.run.x - clipX, dpr),
    lineHeight: 1.2,
    overflow: 'hidden',
    paddingLeft: 6,
    paddingRight: 6,
    position: 'absolute',
    textAlign: input.run.align ?? 'left',
    textDecorationLine: input.run.underline ? 'underline' : input.run.strike ? 'line-through' : undefined,
    MozOsxFontSmoothing: 'grayscale',
    textRendering: 'optimizeLegibility',
    top: visibleClip?.innerTop ?? snapCssPixel(input.run.y - clipY, dpr),
    whiteSpace: input.run.wrap ? 'pre-wrap' : 'pre',
    width: visibleClip?.innerWidth ?? width,
    WebkitFontSmoothing: 'antialiased',
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
  const renderedRuns = useMemo(
    () =>
      active
        ? panes.flatMap((pane) => {
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
              const visibleClip = resolveNativeTextRunVisibleClipV3({ dpr, pane, run, scrollSnapshot: drawScrollSnapshot })
              return visibleClip ? [{ pane, run, visibleClip }] : []
            })
          })
        : [],
    [active, dpr, drawScrollSnapshot, panes],
  )
  const textRunCount = renderedRuns.length

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
      {renderedRuns.map(({ pane, run, visibleClip }) => (
        <div
          data-native-text-run=""
          key={resolveNativeTextRunKey(pane, run)}
          style={resolveNativeTextRunOuterStyleV3({ dpr, pane, run, scrollSnapshot: drawScrollSnapshot, visibleClip })}
        >
          <div style={resolveNativeTextRunInnerStyleV3({ dpr, run, visibleClip })}>{run.text}</div>
        </div>
      ))}
    </div>
  )
})
