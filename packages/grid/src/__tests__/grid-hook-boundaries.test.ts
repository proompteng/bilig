import { describe, expect, test, vi } from 'vitest'
import { collectViewportSubscriptions } from '../useGridViewportSubscriptions.js'
import { canUseWorkerResidentPaneScenes, noteWorkerResidentPaneScenesApplied } from '../useGridSceneResidency.js'
import { resolveRequiresLiveViewportState } from '../useGridSelectionState.js'
import { resolveResizeGuideColumn, resolveResizeGuideRow } from '../useGridResizeState.js'
import { sameBounds } from '../useGridOverlayState.js'
import { visibleRegionFromCamera } from '../useGridCameraState.js'

describe('grid hook boundary helpers', () => {
  test('collects warm and frozen viewport subscriptions without duplicates', () => {
    expect(
      collectViewportSubscriptions({
        viewport: { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
        warmViewports: [{ rowStart: 32, rowEnd: 63, colStart: 256, colEnd: 383 }],
        freezeRows: 1,
        freezeCols: 1,
      }),
    ).toEqual([
      { rowStart: 32, rowEnd: 63, colStart: 256, colEnd: 383 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
      { rowStart: 0, rowEnd: 0, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 0 },
      { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
    ])
  })

  test('keeps worker scenes behind live interaction states without blocking hover overlays', () => {
    const workerResidentPaneScenes = [
      {
        generation: 1,
        paneId: 'body' as const,
        viewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
        surfaceSize: { width: 100, height: 100 },
        gpuScene: { fillRects: [], borderRects: [] },
        textScene: { items: [] },
      },
    ]

    expect(
      canUseWorkerResidentPaneScenes({
        hasActiveHeaderDrag: false,
        hasHoverState: false,
        requiresLiveViewportState: false,
        workerResidentPaneScenes,
      }),
    ).toBe(true)
    expect(
      canUseWorkerResidentPaneScenes({
        hasActiveHeaderDrag: false,
        hasHoverState: true,
        requiresLiveViewportState: false,
        workerResidentPaneScenes,
      }),
    ).toBe(true)
    expect(
      canUseWorkerResidentPaneScenes({
        hasActiveHeaderDrag: true,
        hasHoverState: false,
        requiresLiveViewportState: false,
        workerResidentPaneScenes,
      }),
    ).toBe(false)
  })

  test('notifies perf counters when worker scene packets are applied', () => {
    const noteTypeGpuScenePacketApplied = vi.fn()
    vi.stubGlobal('window', { __biligScrollPerf: { noteTypeGpuScenePacketApplied } })

    noteWorkerResidentPaneScenesApplied([
      {
        generation: 1,
        paneId: 'body',
        viewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
        surfaceSize: { width: 100, height: 100 },
        gpuScene: { fillRects: [], borderRects: [] },
        textScene: { items: [] },
      },
    ])

    expect(noteTypeGpuScenePacketApplied).toHaveBeenCalledWith('body:0:31:0:127')
    vi.unstubAllGlobals()
  })

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
        isRangeMoveDragging: false,
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
        isRangeMoveDragging: false,
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
})
