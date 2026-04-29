import { describe, expect, test } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolveRequiresLiveViewportState } from '../useGridSelectionState.js'
import { resolveResizeGuideColumn, resolveResizeGuideRow } from '../useGridResizeState.js'
import { sameBounds } from '../useGridOverlayState.js'
import { visibleRegionFromCamera } from '../useGridCameraState.js'

describe('grid hook boundary helpers', () => {
  test('resolves live viewport, resize, overlay, and camera helpers', () => {
    expect(
      resolveRequiresLiveViewportState({
        fillPreviewActive: false,
        hasActiveHeaderDrag: false,
        hasActiveResizeColumn: false,
        hasActiveResizeRow: false,
        hasColumnResizePreview: false,
        hasRowResizePreview: true,
        isEditingCell: false,
        isFillHandleDragging: false,
      }),
    ).toBe(true)
    expect(
      resolveRequiresLiveViewportState({
        fillPreviewActive: false,
        hasActiveHeaderDrag: false,
        hasActiveResizeColumn: false,
        hasActiveResizeRow: false,
        hasColumnResizePreview: false,
        hasRowResizePreview: false,
        isEditingCell: true,
        isFillHandleDragging: false,
      }),
    ).toBe(false)
    expect(resolveResizeGuideColumn({ activeResizeColumn: null, cursor: 'col-resize', header: { kind: 'column', index: 4 } })).toBe(4)
    expect(resolveResizeGuideRow({ activeResizeRow: null, cursor: 'row-resize', header: { kind: 'row', index: 6 } })).toBe(6)
    expect(sameBounds({ x: 1, y: 2, width: 3, height: 4 }, { x: 1, y: 2, width: 3, height: 4 })).toBe(true)
    expect(
      visibleRegionFromCamera({
        camera: {
          dpr: 1,
          residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
          scrollLeft: 0,
          scrollTop: 0,
          tx: 7,
          ty: 9,
          updatedAt: 1,
          velocityX: 0,
          velocityY: 0,
          viewportHeight: 100,
          viewportWidth: 100,
          visibleViewport: { rowStart: 2, rowEnd: 5, colStart: 3, colEnd: 8 },
        },
        freezeCols: 1,
        freezeRows: 2,
      }),
    ).toMatchObject({ range: { x: 3, y: 2, width: 6, height: 4 }, tx: 7, ty: 9, freezeRows: 2, freezeCols: 1 })
  })

  test('keeps workbook render hook behind the runtime camera boundary', () => {
    const hookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridRenderState.ts', import.meta.url)), 'utf8')
    const editorHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookEditorOverlayAnchor.ts', import.meta.url)), 'utf8')
    const editorRuntimeHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridEditorRuntime.ts', import.meta.url)), 'utf8')
    const axisRuntimeHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridAxisRuntime.ts', import.meta.url)), 'utf8')
    const geometryRuntimeSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridGeometryRuntime.ts', import.meta.url)), 'utf8')
    const headerHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookHeaderPanes.ts', import.meta.url)), 'utf8')
    const hostRuntimeHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridHostRuntime.ts', import.meta.url)), 'utf8')
    const interactionRuntimeHookSource = readFileSync(
      fileURLToPath(new URL('../useWorkbookGridInteractionRuntime.ts', import.meta.url)),
      'utf8',
    )
    const paneHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridRenderPanes.ts', import.meta.url)), 'utf8')
    const tilePaneHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookRenderTilePanes.ts', import.meta.url)), 'utf8')
    const viewportRuntimeHookSource = readFileSync(fileURLToPath(new URL('../useWorkbookGridViewportRuntime.ts', import.meta.url)), 'utf8')
    const viewportResidencyHookSource = readFileSync(
      fileURLToPath(new URL('../useWorkbookViewportResidencyState.ts', import.meta.url)),
      'utf8',
    )
    const surfaceSource = readFileSync(fileURLToPath(new URL('../WorkbookGridSurface.tsx', import.meta.url)), 'utf8')

    expect(hookSource).toContain("from './useWorkbookGridGeometryRuntime.js'")
    expect(hookSource).not.toContain("from './runtime/gridRuntimeHost.js'")
    expect(geometryRuntimeSource).toContain("from './runtime/gridRuntimeHost.js'")
    expect(hookSource).not.toContain("from './gridCamera.js'")
    expect(hookSource).not.toContain('visibleRegionFromCamera')
    expect(hookSource).not.toContain('scrollCellIntoView')
    expect(hookSource).not.toContain('resolveViewportScrollPosition')
    expect(geometryRuntimeSource).toContain('resolveGridRuntimeGeometryAxes')
    expect(geometryRuntimeSource).not.toContain('Object.entries(')
    expect(geometryRuntimeSource).not.toContain('.toSorted(')
    expect(geometryRuntimeSource).not.toContain('createGridAxisWorldIndexFromRecords')
    expect(geometryRuntimeSource).not.toContain('resolveGridScrollSpacerSize')
    expect(headerHookSource).not.toContain('buildGridGpuScene')
    expect(headerHookSource).not.toContain('buildGridTextScene')
    expect(headerHookSource).not.toContain('buildWorkbookHeaderPaneStatesV3')
    expect(hookSource).not.toContain('useWorkbookHeaderPanes')
    expect(hookSource).not.toContain('useWorkbookRenderTilePanes')
    expect(hookSource).not.toContain('useWorkbookHeaderCellBounds')
    expect(hookSource).not.toContain('useWorkbookViewportResidencyState')
    expect(hookSource).not.toContain('useWorkbookViewportScrollRuntime')
    expect(hookSource).not.toContain('useWorkbookEditorOverlayAnchor')
    expect(hookSource).not.toContain('useWorkbookColumnAutofit')
    expect(hookSource).not.toContain('useWorkbookAxisResizeState')
    expect(hookSource).not.toContain('useWorkbookInteractionOverlayState')
    expect(hookSource).not.toContain('useGridElementSize')
    expect(hookSource).not.toContain('useRef<HTMLDivElement')
    expect(hookSource).not.toContain('useState<VisibleRegionState>')
    expect(paneHookSource).toContain('useWorkbookHeaderPanes')
    expect(paneHookSource).toContain('useWorkbookRenderTilePanes')
    expect(paneHookSource).toContain('useWorkbookHeaderCellBounds')
    expect(viewportRuntimeHookSource).toContain('useWorkbookViewportResidencyState')
    expect(viewportRuntimeHookSource).toContain('useWorkbookViewportScrollRuntime')
    expect(headerHookSource).toContain("from './runtime/gridRuntimeHost.js'")
    expect(headerHookSource).not.toContain("from './runtime/gridHeaderPaneRuntime.js'")
    expect(headerHookSource).not.toContain('useRef<')
    expect(tilePaneHookSource).toContain("from './runtime/gridRuntimeHost.js'")
    expect(tilePaneHookSource).not.toContain("from './runtime/gridRenderTilePaneRuntime.js'")
    expect(tilePaneHookSource).not.toContain('noteRendererTileReadiness')
    expect(tilePaneHookSource).not.toContain('.subscribeCells(')
    expect(tilePaneHookSource).not.toContain('clearRetainedRenderTilePanes')
    expect(tilePaneHookSource).not.toContain('useRef<')
    expect(tilePaneHookSource).not.toContain('buildViewportTileInterest')
    expect(tilePaneHookSource).not.toContain('.subscribeRenderTileDeltas(')
    expect(hookSource).not.toContain('projectionViewportRevision')
    expect(hookSource).not.toContain('setProjectionViewportRevision')
    expect(tilePaneHookSource).toContain('subscribeViewport')
    expect(tilePaneHookSource).toContain('noteProjectedViewportPatch')
    expect(viewportResidencyHookSource).toContain("from './runtime/gridRuntimeHost.js'")
    expect(viewportResidencyHookSource).not.toContain("from './runtime/gridViewportResidencyRuntime.js'")
    expect(viewportResidencyHookSource).not.toContain('.subscribeCells(')
    expect(viewportResidencyHookSource).not.toContain('setSceneRevision')
    expect(editorHookSource).toContain("from './runtime/gridEditorAnchorRuntime.js'")
    expect(editorHookSource).not.toContain('useRef<GridEditorAnchorRuntime')
    expect(editorHookSource).not.toContain('resolveEditorOverlayScreenBounds')
    expect(editorHookSource).not.toContain('applyEditorOverlayBounds')
    expect(editorHookSource).not.toContain('snapshotToRenderCell')
    expect(editorRuntimeHookSource).toContain('useWorkbookEditorOverlayAnchor')
    expect(editorRuntimeHookSource).toContain('useWorkbookColumnAutofit')
    expect(axisRuntimeHookSource).toContain('useWorkbookAxisResizeState')
    expect(interactionRuntimeHookSource).toContain('useWorkbookInteractionOverlayState')
    expect(hostRuntimeHookSource).toContain('useGridElementSize')
    expect(hostRuntimeHookSource).toContain('useRef<HTMLDivElement')
    expect(hostRuntimeHookSource).toContain('useState<VisibleRegionState>')
    expect(surfaceSource).not.toContain('createGridGeometrySnapshot')
  })
})
