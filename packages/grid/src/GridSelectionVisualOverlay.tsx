import { useEffect, useMemo, useState, type CSSProperties } from 'react'
import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import {
  buildGridSelectionVisualRects,
  type GridSelectionVisualRect,
  type GridSelectionVisualRectRole,
} from './gridSelectionVisualRects.js'
import type { WorkbookGridScrollStore } from './workbookGridScrollStore.js'
import { workbookThemeColors } from './workbookTheme.js'

export { buildGridSelectionVisualRects } from './gridSelectionVisualRects.js'
export type { GridSelectionVisualRect, GridSelectionVisualRectRole } from './gridSelectionVisualRects.js'

const SELECTION_STROKE_WIDTH = 1
const SELECTION_BORDER_EDGE_NAMES = ['top', 'right', 'bottom', 'left'] as const

export interface GridSelectionVisualOverlayProps {
  readonly geometry?: GridGeometrySnapshot | null | undefined
  readonly getGeometrySnapshot?: (() => GridGeometrySnapshot | null) | undefined
  readonly gridSelection: GridSelection
  readonly hoverCell?: Item | null | undefined
  readonly selectedCell: Item
  readonly selectionChromeMode?: 'visible' | 'geometry-only' | 'chrome-only' | undefined
  readonly selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  readonly showFillHandle: boolean
  readonly scrollTransformStore?: WorkbookGridScrollStore | undefined
}

export function GridSelectionVisualOverlay(props: GridSelectionVisualOverlayProps) {
  const {
    geometry: staticGeometry,
    getGeometrySnapshot,
    gridSelection,
    hoverCell,
    scrollTransformStore,
    selectedCell,
    selectionChromeMode = 'visible',
    selectionRange,
    showFillHandle,
  } = props
  const [scrollVersion, setScrollVersion] = useState(0)
  useEffect(() => {
    if (!scrollTransformStore) {
      return
    }
    return scrollTransformStore.subscribe(() => {
      setScrollVersion((version) => version + 1)
    })
  }, [scrollTransformStore])
  const geometry = useMemo(() => {
    void scrollVersion
    return getGeometrySnapshot?.() ?? staticGeometry ?? null
  }, [getGeometrySnapshot, staticGeometry, scrollVersion])
  const rects = useMemo(
    () =>
      geometry
        ? buildGridSelectionVisualRects({
            geometry,
            gridSelection,
            hoverCell: hoverCell ?? null,
            selectedCell,
            selectionRange,
            showFillHandle,
          })
        : [],
    [geometry, gridSelection, hoverCell, selectedCell, selectionRange, showFillHandle],
  )

  if (rects.length === 0) {
    return null
  }

  return (
    <div aria-hidden="true" className="pointer-events-none absolute inset-0 overflow-hidden" data-testid="grid-selection-visual-overlay">
      {rects.map((rect) => (
        <div
          className={classNameForRole(rect.role)}
          data-grid-selection-visual-key={rect.key}
          data-grid-selection-visual-role={rect.role}
          key={keyForRect(rect)}
          style={styleForRect(rect, selectionChromeMode)}
        >
          {rect.role === 'selection-border' || rect.role === 'active-border'
            ? SELECTION_BORDER_EDGE_NAMES.map((edge) => (
                <div aria-hidden="true" data-grid-selection-visual-edge={edge} key={edge} style={styleForSelectionBorderEdge(rect, edge)} />
              ))
            : null}
        </div>
      ))}
    </div>
  )
}

function classNameForRole(role: GridSelectionVisualRectRole): string {
  switch (role) {
    case 'selection-fill':
      return 'absolute box-border'
    case 'selection-gridline':
      return 'absolute box-border'
    case 'header-fill':
      return 'absolute box-border'
    case 'header-seam-cover':
      return 'absolute box-border'
    case 'hover-fill':
      return 'absolute box-border'
    case 'selection-border':
      return 'absolute box-border'
    case 'active-border':
      return 'absolute box-border'
    case 'fill-handle':
      return 'absolute box-border rounded-[1px] border border-[var(--wb-surface)] bg-[var(--wb-accent)]'
  }
}

function keyForRect(rect: GridSelectionVisualRect): string {
  return `${rect.role}:${rect.key}`
}

function styleForRect(
  rect: GridSelectionVisualRect,
  mode: NonNullable<GridSelectionVisualOverlayProps['selectionChromeMode']>,
): CSSProperties {
  const hidden = mode === 'geometry-only' || (mode === 'chrome-only' && rect.role === 'hover-fill')
  const base = {
    height: rect.bounds.height,
    left: rect.bounds.x,
    opacity: hidden ? 0 : undefined,
    top: rect.bounds.y,
    width: rect.bounds.width,
    zIndex: rect.role === 'fill-handle' ? 50 : 20,
  }
  if (rect.role === 'selection-border' || rect.role === 'active-border') {
    return {
      ...base,
      backgroundColor: 'transparent',
      boxSizing: 'border-box',
      overflow: 'visible',
      boxShadow: 'none',
      outline: 'none',
    }
  }
  if (rect.role === 'selection-gridline') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.gridBorder,
    }
  }
  if (rect.role === 'fill-handle') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.selectionAccent,
      borderColor: workbookThemeColors.surface,
      borderWidth: 1,
      boxSizing: 'border-box',
    }
  }
  if (rect.role === 'header-fill') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.selectionHeaderOverlayFill,
    }
  }
  if (rect.role === 'header-seam-cover') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.selectionHeaderSeamFill,
    }
  }
  if (rect.role === 'hover-fill') {
    return {
      ...base,
      backgroundColor: workbookThemeColors.hoverFill,
    }
  }
  return {
    ...base,
    backgroundColor: workbookThemeColors.selectionFill,
  }
}

function styleForSelectionBorderEdge(rect: GridSelectionVisualRect, edge: (typeof SELECTION_BORDER_EDGE_NAMES)[number]): CSSProperties {
  const strokeWidth = Math.max(1, rect.strokeWidth ?? SELECTION_STROKE_WIDTH)
  const base: CSSProperties = {
    backgroundColor: workbookThemeColors.selectionAccent,
    position: 'absolute',
  }
  switch (edge) {
    case 'top':
      return {
        ...base,
        height: strokeWidth,
        left: -strokeWidth,
        right: 0,
        top: -strokeWidth,
      }
    case 'right':
      return {
        ...base,
        bottom: strokeWidth,
        right: 0,
        top: 0,
        width: strokeWidth,
      }
    case 'bottom':
      return {
        ...base,
        bottom: 0,
        height: strokeWidth,
        left: -strokeWidth,
        right: 0,
      }
    case 'left':
      return {
        ...base,
        bottom: strokeWidth,
        left: -strokeWidth,
        top: 0,
        width: strokeWidth,
      }
  }
}
