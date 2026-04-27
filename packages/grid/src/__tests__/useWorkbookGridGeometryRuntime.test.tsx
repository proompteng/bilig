// @vitest-environment jsdom
import { act, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { useWorkbookGridGeometryRuntime, type WorkbookGridGeometryRuntimeState } from '../useWorkbookGridGeometryRuntime.js'

describe('useWorkbookGridGeometryRuntime', () => {
  it('owns axis indexes, runtime stores, and cell geometry outside the render hook', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    let latestState: WorkbookGridGeometryRuntimeState | null = null
    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const scrollViewportRef = useRef<HTMLDivElement | null>(null)
      latestState = useWorkbookGridGeometryRuntime({
        columnWidths: { 0: 150, 3: 140 },
        controlledHiddenColumns: { 4: true },
        controlledHiddenRows: { 6: true },
        freezeCols: 1,
        freezeRows: 1,
        gridMetrics,
        hostClientHeight: 240,
        hostClientWidth: 480,
        hostRef,
        rowHeights: { 0: 30, 2: 40 },
        scrollViewportRef,
        sheetName: 'Sheet1',
      })
      return (
        <div
          ref={(node) => {
            hostRef.current = node
            if (node) {
              node.getBoundingClientRect = () =>
                ({
                  bottom: 280,
                  height: 240,
                  left: 10,
                  right: 490,
                  top: 40,
                  width: 480,
                  x: 10,
                  y: 40,
                  toJSON: () => ({}),
                }) as DOMRect
            }
          }}
        >
          <div
            ref={(node) => {
              scrollViewportRef.current = node
              if (node) {
                Object.defineProperty(node, 'clientWidth', { configurable: true, value: 480 })
                Object.defineProperty(node, 'clientHeight', { configurable: true, value: 240 })
                node.scrollLeft = 25
                node.scrollTop = 15
              }
            }}
          />
        </div>
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestState?.columnWidthOverridesAttr).toBe('{"0":150,"3":140}')
    expect(latestState?.rowHeightOverridesAttr).toBe('{"0":30,"2":40}')
    expect(latestState?.frozenColumnWidth).toBe(150)
    expect(latestState?.frozenRowHeight).toBe(30)
    expect(latestState?.getCellLocalBounds(4, 1)).toBeUndefined()
    expect(latestState?.getCellLocalBounds(1, 6)).toBeUndefined()

    latestState!.scrollTransformRef.current = {
      scrollLeft: 25,
      scrollTop: 15,
      tx: 0,
      ty: 0,
    }
    expect(latestState?.getCellLocalBounds(1, 1)).toEqual({
      height: gridMetrics.rowHeight,
      width: gridMetrics.columnWidth,
      x: gridMetrics.rowMarkerWidth + 150 - 25,
      y: gridMetrics.headerHeight + 30 - 15,
    })
    expect(latestState?.getCellScreenBounds(1, 1)).toEqual({
      height: gridMetrics.rowHeight,
      width: gridMetrics.columnWidth,
      x: 10 + gridMetrics.rowMarkerWidth + 150 - 25,
      y: 40 + gridMetrics.headerHeight + 30 - 15,
    })
    expect(latestState?.getLiveGeometrySnapshot()?.camera).toMatchObject({
      bodyViewportHeight: expect.any(Number),
      bodyViewportWidth: expect.any(Number),
      scrollLeft: 25,
      scrollTop: 15,
      sheetName: 'Sheet1',
    })

    await act(async () => {
      root.unmount()
    })
  })
})
