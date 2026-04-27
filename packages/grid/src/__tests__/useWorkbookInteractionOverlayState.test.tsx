// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it } from 'vitest'
import { createGridSelection } from '../gridSelection.js'
import type { Rectangle } from '../gridTypes.js'
import { useWorkbookInteractionOverlayState, type WorkbookInteractionOverlayState } from '../useWorkbookInteractionOverlayState.js'

describe('useWorkbookInteractionOverlayState', () => {
  it('syncs simple selection to the selected cell and preserves explicit range selections', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let selectedCol = 1
    let selectedRow = 2
    let latestState: WorkbookInteractionOverlayState | null = null

    function Harness() {
      latestState = useWorkbookInteractionOverlayState({
        activeResizeColumn: null,
        activeResizeRow: null,
        getCellLocalBounds: () => undefined,
        hasColumnResizePreview: false,
        hasRowResizePreview: false,
        isEditingCell: false,
        selectedCol,
        selectedRow,
        visibleRange: { height: 20, width: 10, x: 0, y: 0 },
      })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.gridSelection.current?.cell).toEqual([1, 2])

    selectedCol = 3
    selectedRow = 4
    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.gridSelection.current?.cell).toEqual([3, 4])

    await act(async () => {
      latestState?.setGridSelection({
        ...createGridSelection(3, 4),
        current: { cell: [3, 4], range: { height: 2, width: 2, x: 3, y: 4 }, rangeStack: [] },
      })
    })
    selectedCol = 6
    selectedRow = 7
    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.gridSelection.current?.range).toEqual({ height: 2, width: 2, x: 3, y: 4 })

    await act(async () => {
      root.unmount()
    })
  })

  it('computes fill-preview bounds and live viewport state for visual interactions', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const cellBounds = new Map<string, Rectangle>([
      ['1:2', { height: 22, width: 104, x: 120, y: 64 }],
      ['2:2', { height: 22, width: 104, x: 224, y: 64 }],
    ])
    let latestState: WorkbookInteractionOverlayState | null = null

    function Harness() {
      latestState = useWorkbookInteractionOverlayState({
        activeResizeColumn: null,
        activeResizeRow: null,
        getCellLocalBounds: (col, row) => cellBounds.get(`${col}:${row}`),
        hasColumnResizePreview: false,
        hasRowResizePreview: false,
        isEditingCell: false,
        selectedCol: 1,
        selectedRow: 2,
        visibleRange: { height: 20, width: 10, x: 0, y: 0 },
      })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.requiresLiveViewportState).toBe(false)

    await act(async () => {
      latestState?.setFillPreviewRange({ height: 1, width: 2, x: 1, y: 2 })
    })

    expect(latestState?.requiresLiveViewportState).toBe(true)
    expect(latestState?.fillPreviewBounds).toEqual({ height: 22, width: 208, x: 120, y: 64 })

    await act(async () => {
      root.unmount()
    })
  })
})
