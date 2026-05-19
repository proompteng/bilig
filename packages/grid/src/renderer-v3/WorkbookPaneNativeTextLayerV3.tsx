import { memo, useMemo, useSyncExternalStore, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'
import { workbookNativeTextQualityStyle } from '../workbookTextQuality.js'
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

export interface SuppressedNativeTextCellV3 {
  readonly col: number
  readonly row: number
}

export interface WorkbookPaneNativeTextLayerV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly presentedScrollSnapshot?: WorkbookGridScrollSnapshot | null | undefined
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly suppressedTextCell?: SuppressedNativeTextCellV3 | null | undefined
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

export function resolveNativeTextLayerDrawScrollSnapshotV3(input: {
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

function isPaneDrawVisible(pane: TextLayerPane): boolean {
  return pane.drawVisible !== false
}

function getPaneTextRuns(pane: TextLayerPane): readonly TextQuadRun[] {
  return 'tile' in pane ? pane.tile.textRuns : pane.textRuns
}

function isNativeTextRunSuppressed(
  pane: TextLayerPane,
  run: TextQuadRun,
  suppressedTextCell: SuppressedNativeTextCellV3 | null | undefined,
): boolean {
  return Boolean(suppressedTextCell && 'tile' in pane && run.row === suppressedTextCell.row && run.col === suppressedTextCell.col)
}

function resolveNativeTextRunKey(pane: TextLayerPane, run: TextQuadRun): string {
  const paneIdentity =
    'tile' in pane
      ? [pane.paneId, pane.tile.tileId, pane.tile.coord.sheetOrdinal, pane.tile.coord.rowTile, pane.tile.coord.colTile].join(':')
      : pane.paneId
  const runIdentity =
    run.row !== undefined && run.col !== undefined ? ['cell', run.row, run.col].join(':') : ['header', run.text, run.x, run.y].join(':')
  return [
    paneIdentity,
    runIdentity,
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

function subscribeDevicePixelRatioChange(onStoreChange: () => void): () => void {
  if (typeof window === 'undefined') {
    return subscribeNoop()
  }

  let disposed = false
  let resolutionQuery: MediaQueryList | null = null
  const handleChange = () => {
    if (disposed) {
      return
    }
    onStoreChange()
    resetResolutionQuery()
  }
  const resetResolutionQuery = () => {
    if (disposed) {
      return
    }
    removeResolutionQueryListener(resolutionQuery, handleChange)
    resolutionQuery = typeof window.matchMedia === 'function' ? window.matchMedia(`(resolution: ${getDevicePixelRatio()}dppx)`) : null
    addResolutionQueryListener(resolutionQuery, handleChange)
  }

  resetResolutionQuery()
  window.addEventListener('resize', handleChange)
  window.visualViewport?.addEventListener('resize', handleChange)
  return () => {
    disposed = true
    removeResolutionQueryListener(resolutionQuery, handleChange)
    resolutionQuery = null
    window.removeEventListener('resize', handleChange)
    window.visualViewport?.removeEventListener('resize', handleChange)
  }
}

function addResolutionQueryListener(query: MediaQueryList | null, listener: () => void): void {
  if (!query) {
    return
  }
  if (typeof query.addEventListener === 'function') {
    query.addEventListener('change', listener)
    return
  }
  const legacyQuery = query as MediaQueryList & { addListener?: (listener: () => void) => void }
  legacyQuery.addListener?.(listener)
}

function removeResolutionQueryListener(query: MediaQueryList | null, listener: () => void): void {
  if (!query) {
    return
  }
  if (typeof query.removeEventListener === 'function') {
    query.removeEventListener('change', listener)
    return
  }
  const legacyQuery = query as MediaQueryList & { removeListener?: (listener: () => void) => void }
  legacyQuery.removeListener?.(listener)
}

function snapCssPixel(value: number, dpr: number): number {
  return Math.round(value * dpr) / dpr
}

function resolveNativeTextLineBoxV3(input: { readonly run: TextQuadRun; readonly dpr: number }): {
  readonly height: number
  readonly topInset: number
} {
  const fontStyle = resolveNativeTextRunFontStyleV3(input.run)
  const contentHeight = input.run.height ?? 0
  const lineHeight = snapCssPixel(fontStyle.fontSize * 1.2, input.dpr)
  return {
    height: lineHeight,
    topInset: snapCssPixel(Math.max(0, (contentHeight - lineHeight) / 2), input.dpr),
  }
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
  const lineBox = resolveNativeTextLineBoxV3({ dpr, run: input.run })
  const baseTop = visibleClip?.innerTop ?? snapCssPixel(input.run.y - clipY, dpr)
  const textTop = input.run.wrap ? baseTop : snapCssPixel(baseTop + lineBox.topInset, dpr)
  return {
    ...workbookNativeTextQualityStyle,
    boxSizing: 'border-box',
    color: input.run.color ?? '#111827',
    display: 'block',
    fontFamily: fontStyle.fontFamily,
    fontFeatureSettings: 'normal',
    fontSize: fontStyle.fontSize,
    fontStyle: fontStyle.fontStyle,
    fontVariantNumeric: 'tabular-nums',
    fontWeight: fontStyle.fontWeight,
    height: input.run.wrap ? (visibleClip?.innerHeight ?? height) : lineBox.height,
    letterSpacing: 0,
    left: visibleClip?.innerLeft ?? snapCssPixel(input.run.x - clipX, dpr),
    lineHeight: `${lineBox.height}px`,
    overflow: 'hidden',
    paddingLeft: 6,
    paddingRight: 6,
    position: 'absolute',
    textAlign: input.run.align ?? 'left',
    textDecorationLine: input.run.underline ? 'underline' : input.run.strike ? 'line-through' : undefined,
    top: textTop,
    whiteSpace: input.run.wrap ? 'pre-wrap' : 'pre',
    width: visibleClip?.innerWidth ?? width,
  }
}

export const WorkbookPaneNativeTextLayerV3 = memo(function WorkbookPaneNativeTextLayerV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes,
  presentedScrollSnapshot = null,
  scrollTransformStore,
  suppressedTextCell = null,
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
      resolveNativeTextLayerDrawScrollSnapshotV3({
        geometry: resolvedGeometry,
        liveScrollSnapshot: scrollSnapshot,
        panes: tilePanes,
        presentedScrollSnapshot,
      }),
    [presentedScrollSnapshot, resolvedGeometry, scrollSnapshot, tilePanes],
  )
  const panes = useMemo<readonly TextLayerPane[]>(() => [...tilePanes, ...headerPanes], [headerPanes, tilePanes])
  const dpr = useSyncExternalStore(subscribeDevicePixelRatioChange, getDevicePixelRatio, () => 1)
  const renderedRuns = useMemo(
    () =>
      active
        ? panes.flatMap((pane) => {
            if (!isPaneDrawVisible(pane)) {
              return []
            }
            return getPaneTextRuns(pane).flatMap((run) => {
              if (isNativeTextRunSuppressed(pane, run, suppressedTextCell)) {
                return []
              }
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
    [active, dpr, drawScrollSnapshot, panes, suppressedTextCell],
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
      data-v3-native-text-presented-scroll-left={drawScrollSnapshot.scrollLeft ?? ''}
      data-v3-native-text-presented-scroll-top={drawScrollSnapshot.scrollTop ?? ''}
      data-v3-native-text-render-tx={drawScrollSnapshot.renderTx ?? drawScrollSnapshot.tx}
      data-v3-native-text-render-ty={drawScrollSnapshot.renderTy ?? drawScrollSnapshot.ty}
      data-v3-native-text-run-count={textRunCount}
      data-testid="grid-native-text-layer"
      style={{ contain: 'strict' }}
    >
      {renderedRuns.map(({ pane, run, visibleClip }) => (
        <div
          data-native-text-run=""
          data-native-text-run-col={run.col ?? ''}
          data-native-text-run-row={run.row ?? ''}
          key={resolveNativeTextRunKey(pane, run)}
          style={resolveNativeTextRunOuterStyleV3({ dpr, pane, run, scrollSnapshot: drawScrollSnapshot, visibleClip })}
        >
          <div style={resolveNativeTextRunInnerStyleV3({ dpr, run, visibleClip })}>{run.text}</div>
        </div>
      ))}
    </div>
  )
})
