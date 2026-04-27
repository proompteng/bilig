// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, MIN_COLUMN_WIDTH, MIN_ROW_HEIGHT } from '../gridMetrics.js'
import { useWorkbookAxisResizeState, type WorkbookAxisResizeState } from '../useWorkbookAxisResizeState.js'

describe('useWorkbookAxisResizeState', () => {
  it('keeps resize preview state out of the render hook and clamps committed axis sizes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let latestState: WorkbookAxisResizeState | null = null

    function Harness() {
      latestState = useWorkbookAxisResizeState({ sheetName: 'Sheet1' })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    await act(async () => {
      expect(latestState?.previewColumnWidth(2, MAX_COLUMN_WIDTH + 1000)).toBe(MAX_COLUMN_WIDTH)
      expect(latestState?.previewRowHeight(3, MAX_ROW_HEIGHT + 1000)).toBe(MAX_ROW_HEIGHT)
    })

    expect(latestState?.hasColumnResizePreview).toBe(true)
    expect(latestState?.hasRowResizePreview).toBe(true)
    expect(latestState?.columnWidths[2]).toBe(MAX_COLUMN_WIDTH)
    expect(latestState?.rowHeights[3]).toBe(MAX_ROW_HEIGHT)
    expect(latestState?.getPreviewColumnWidth(2)).toBe(MAX_COLUMN_WIDTH)
    expect(latestState?.getPreviewRowHeight(3)).toBe(MAX_ROW_HEIGHT)

    await act(async () => {
      latestState?.clearColumnResizePreview(2)
      latestState?.clearRowResizePreview(3)
    })

    expect(latestState?.hasColumnResizePreview).toBe(false)
    expect(latestState?.hasRowResizePreview).toBe(false)
    expect(latestState?.columnWidths[2]).toBeUndefined()
    expect(latestState?.rowHeights[3]).toBeUndefined()

    await act(async () => {
      latestState?.commitColumnWidth(2, MIN_COLUMN_WIDTH - 1000)
      latestState?.commitRowHeight(3, MIN_ROW_HEIGHT - 1000)
    })

    expect(latestState?.columnWidths[2]).toBe(MIN_COLUMN_WIDTH)
    expect(latestState?.rowHeights[3]).toBe(MIN_ROW_HEIGHT)

    await act(async () => {
      root.unmount()
    })
  })

  it('delegates controlled commits and applies hidden-axis overrides', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onColumnWidthChange = vi.fn()
    const onRowHeightChange = vi.fn()
    let latestState: WorkbookAxisResizeState | null = null

    function Harness() {
      latestState = useWorkbookAxisResizeState({
        controlledColumnWidths: { 4: 120 },
        controlledHiddenColumns: { 5: true },
        controlledHiddenRows: { 7: true },
        controlledRowHeights: { 6: 32 },
        onColumnWidthChange,
        onRowHeightChange,
        sheetName: 'Sheet1',
      })
      return null
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.columnWidths[4]).toBe(120)
    expect(latestState?.columnWidths[5]).toBe(0)
    expect(latestState?.rowHeights[6]).toBe(32)
    expect(latestState?.rowHeights[7]).toBe(0)

    await act(async () => {
      latestState?.commitColumnWidth(4, MAX_COLUMN_WIDTH + 1)
      latestState?.commitRowHeight(6, MIN_ROW_HEIGHT - 1)
    })

    expect(onColumnWidthChange).toHaveBeenCalledWith(4, MAX_COLUMN_WIDTH)
    expect(onRowHeightChange).toHaveBeenCalledWith(6, MIN_ROW_HEIGHT)

    await act(async () => {
      root.unmount()
    })
  })
})
