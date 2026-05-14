import { memo, useSyncExternalStore } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'
import type { GridHeaderPaneState } from '../gridHeaderPanes.js'
import type { GridCameraStore } from '../runtime/gridCameraStore.js'
import { WORKBOOK_FONT_SANS } from '../workbookTheme.js'
import type { WorkbookGridScrollSnapshot, WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import type { TextQuadRun } from './line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from './render-tile-pane-state.js'
import { resolveTypeGpuV3DrawScrollSnapshot, resolveWorkbookPaneRendererGeometryV3 } from './workbook-pane-renderer-runtime.js'

type TextPane = GridHeaderPaneState | WorkbookRenderTilePaneState

export interface WorkbookPaneTextOverlayV3Props {
  readonly active: boolean
  readonly cameraStore?: GridCameraStore | null | undefined
  readonly geometry: GridGeometrySnapshot | null
  readonly headerPanes: readonly GridHeaderPaneState[]
  readonly host: HTMLDivElement | null
  readonly scrollTransformStore: WorkbookGridScrollStore | null
  readonly tilePanes: readonly WorkbookRenderTilePaneState[]
}

const EMPTY_SCROLL_SNAPSHOT: WorkbookGridScrollSnapshot = Object.freeze({ tx: 0, ty: 0 })

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

export const WorkbookPaneTextOverlayV3 = memo(function WorkbookPaneTextOverlayV3({
  active,
  cameraStore = null,
  geometry,
  headerPanes,
  host,
  scrollTransformStore,
  tilePanes,
}: WorkbookPaneTextOverlayV3Props) {
  const subscribeCamera = cameraStore ? (listener: () => void) => cameraStore.subscribe(listener) : noopSubscribe
  const getCameraSnapshot = cameraStore ? () => cameraStore.getSnapshot() : getNullGeometrySnapshot
  const subscribeScroll = scrollTransformStore ? (listener: () => void) => scrollTransformStore.subscribe(listener) : noopSubscribe
  const getScrollSnapshot = scrollTransformStore ? () => scrollTransformStore.getSnapshot() : getEmptyScrollSnapshot
  const cameraGeometry = useSyncExternalStore(subscribeCamera, getCameraSnapshot, getNullGeometrySnapshot)
  const scrollSnapshot = useSyncExternalStore(subscribeScroll, getScrollSnapshot, getEmptyScrollSnapshot)

  if (!active || !host) {
    return null
  }

  const resolvedGeometry = resolveWorkbookPaneRendererGeometryV3({
    cameraStore: null,
    geometry: cameraGeometry ?? geometry,
  })
  const drawScrollSnapshot = resolveTypeGpuV3DrawScrollSnapshot({
    fallback: scrollSnapshot,
    geometry: resolvedGeometry,
    panes: tilePanes,
  })
  const panes: readonly TextPane[] = [...tilePanes, ...headerPanes]

  return (
    <div
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-[11] overflow-hidden"
      data-testid="grid-pane-text-overlay"
      style={{ contain: 'strict' }}
    >
      {panes.map((pane) => renderPaneText(pane, drawScrollSnapshot))}
    </div>
  )
})

function renderPaneText(pane: TextPane, scrollSnapshot: WorkbookGridScrollSnapshot) {
  if (pane.frame.width <= 0 || pane.frame.height <= 0) {
    return null
  }
  const textRuns = 'tile' in pane ? pane.tile.textRuns : pane.textRuns
  if (textRuns.length === 0) {
    return null
  }
  const offset = resolvePaneRenderOffset(pane, scrollSnapshot)
  return (
    <div
      key={pane.paneId}
      className="absolute overflow-hidden"
      style={{
        height: pane.frame.height,
        left: pane.frame.x,
        top: pane.frame.y,
        width: pane.frame.width,
      }}
    >
      {textRuns.map((run, index) => renderTextRun(run, offset, `${pane.paneId}:${index}`))}
    </div>
  )
}

function renderTextRun(run: TextQuadRun, offset: { readonly x: number; readonly y: number }, key: string) {
  const runWidth = run.width ?? 0
  const runHeight = run.height ?? 0
  const clipX = run.clipX ?? run.x
  const clipY = run.clipY ?? run.y
  const clipWidth = run.clipWidth ?? runWidth
  const clipHeight = run.clipHeight ?? runHeight
  if (!run.text || clipWidth <= 0 || clipHeight <= 0) {
    return null
  }
  const fontSize = run.fontSize ?? 12
  const font = run.font?.trim() ? run.font : `400 ${fontSize}px ${WORKBOOK_FONT_SANS}`
  const textDecorationLine = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined
  return (
    <span
      key={key}
      className="absolute block overflow-hidden whitespace-nowrap"
      style={{
        color: run.color ?? '#111827',
        font,
        height: clipHeight,
        left: offset.x + clipX,
        lineHeight: `${runHeight}px`,
        textAlign: run.align ?? 'left',
        textDecorationLine,
        top: offset.y + clipY,
        width: clipWidth,
      }}
    >
      <span
        className="absolute block overflow-hidden whitespace-nowrap"
        style={{
          height: runHeight,
          left: run.x - clipX,
          paddingLeft: run.align === 'left' || !run.align ? 6 : 0,
          paddingRight: run.align === 'right' ? 6 : 0,
          top: run.y - clipY,
          width: runWidth,
        }}
      >
        {run.text}
      </span>
    </span>
  )
}

function noopSubscribe(): () => void {
  return () => {}
}

function getNullGeometrySnapshot(): GridGeometrySnapshot | null {
  return null
}

function getEmptyScrollSnapshot(): WorkbookGridScrollSnapshot {
  return EMPTY_SCROLL_SNAPSHOT
}
