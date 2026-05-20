// @vitest-environment jsdom
import { act, useMemo, useRef } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, test } from 'vitest'
import { createGridAxisWorldIndex } from '../gridAxisWorldIndex.js'
import { createGridGeometrySnapshotFromAxes } from '../gridGeometry.js'
import { getGridMetrics } from '../gridMetrics.js'
import { createGridSelection } from '../gridSelection.js'
import { useWorkbookGridPointerResolvers } from '../useWorkbookGridPointerResolvers.js'

describe('useWorkbookGridPointerResolvers', () => {
  afterEach(() => {
    document.body.innerHTML = ''
  })

  test('maps body clicks through the current live scrolled camera', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    let latestResolvers: ReturnType<typeof useWorkbookGridPointerResolvers> | null = null

    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const liveGeometry = useMemo(
        () =>
          createGridGeometrySnapshotFromAxes({
            columns: createGridAxisWorldIndex({ axisLength: 300, defaultSize: gridMetrics.columnWidth }),
            dpr: 1,
            gridMetrics,
            hostHeight: 420,
            hostWidth: 640,
            rows: createGridAxisWorldIndex({ axisLength: 300, defaultSize: gridMetrics.rowHeight }),
            scrollLeft: gridMetrics.columnWidth,
            scrollTop: gridMetrics.rowHeight * 150,
            sheetName: 'Sheet1',
            updatedAt: 100,
          }),
        [],
      )
      latestResolvers = useWorkbookGridPointerResolvers({
        getGeometrySnapshot: () => liveGeometry,
        gridSelection: createGridSelection(0, 0),
        hostRef,
        selectedCell: { col: 0, row: 0 },
      })

      return (
        <div
          ref={(node) => {
            hostRef.current = node
            if (node) {
              node.getBoundingClientRect = () =>
                ({
                  bottom: 420,
                  height: 420,
                  left: 20,
                  right: 660,
                  top: 10,
                  width: 640,
                  x: 20,
                  y: 10,
                  toJSON: () => ({}),
                }) as DOMRect
            }
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(
      latestResolvers?.resolvePointerCell(
        20 + gridMetrics.rowMarkerWidth + gridMetrics.columnWidth * 2 + 1,
        10 + gridMetrics.headerHeight + gridMetrics.rowHeight * 3 + 1,
      ),
    ).toEqual([3, 153])

    await act(async () => {
      root.unmount()
    })
  })

  test('uses live camera geometry without falling back to stale visible-region math', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    const staleVisibleRegion = {
      freezeCols: 0,
      freezeRows: 0,
      range: { x: 0, y: 0, width: 12, height: 24 },
      tx: 0,
      ty: 0,
    }
    let latestResolvers: ReturnType<typeof useWorkbookGridPointerResolvers> | null = null

    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const liveGeometry = useMemo(
        () =>
          createGridGeometrySnapshotFromAxes({
            columns: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.columnWidth }),
            dpr: 1,
            gridMetrics,
            hostHeight: 180,
            hostWidth: 220,
            rows: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.rowHeight }),
            scrollLeft: 0,
            scrollTop: 0,
            sheetName: 'Sheet1',
            updatedAt: 100,
          }),
        [],
      )
      latestResolvers = useWorkbookGridPointerResolvers({
        getGeometrySnapshot: () => liveGeometry,
        gridSelection: createGridSelection(0, 0),
        hostRef,
        selectedCell: { col: 0, row: 0 },
      })

      return (
        <div
          ref={(node) => {
            hostRef.current = node
            if (node) {
              node.getBoundingClientRect = () =>
                ({
                  bottom: 240,
                  height: 240,
                  left: 0,
                  right: 1000,
                  top: 0,
                  width: 1000,
                  x: 0,
                  y: 0,
                  toJSON: () => ({}),
                }) as DOMRect
            }
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(latestResolvers?.resolveHeaderSelectionAtPointer(300, 10, staleVisibleRegion)).toBeNull()
    expect(latestResolvers?.resolvePointerCell(300, 80, staleVisibleRegion)).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  test('does not let a clipped selected cell steal header clicks', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    let latestResolvers: ReturnType<typeof useWorkbookGridPointerResolvers> | null = null

    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const liveGeometry = useMemo(
        () =>
          createGridGeometrySnapshotFromAxes({
            columns: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.columnWidth }),
            dpr: 1,
            gridMetrics,
            hostHeight: 180,
            hostWidth: 420,
            rows: createGridAxisWorldIndex({ axisLength: 100, defaultSize: gridMetrics.rowHeight }),
            scrollLeft: 0,
            scrollTop: 10,
            sheetName: 'Sheet1',
            updatedAt: 100,
          }),
        [],
      )
      latestResolvers = useWorkbookGridPointerResolvers({
        getGeometrySnapshot: () => liveGeometry,
        gridSelection: createGridSelection(1, 0),
        hostRef,
        selectedCell: { col: 1, row: 0 },
      })

      return (
        <div
          ref={(node) => {
            hostRef.current = node
            if (node) {
              node.getBoundingClientRect = () =>
                ({
                  bottom: 190,
                  height: 180,
                  left: 10,
                  right: 430,
                  top: 10,
                  width: 420,
                  x: 10,
                  y: 10,
                  toJSON: () => ({}),
                }) as DOMRect
            }
          }}
        />
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    const headerClientX = 10 + gridMetrics.rowMarkerWidth + gridMetrics.columnWidth + 20
    const headerClientY = 10 + gridMetrics.headerHeight - 2
    expect(latestResolvers?.resolvePointerCell(headerClientX, headerClientY)).toBeNull()
    expect(latestResolvers?.resolveHeaderSelectionAtPointer(headerClientX, headerClientY)).toEqual({ kind: 'column', index: 1 })

    await act(async () => {
      root.unmount()
    })
  })

  test('resolves header selections using host coordinates', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const gridMetrics = getGridMetrics()
    let latestResolvers: ReturnType<typeof useWorkbookGridPointerResolvers> | null = null

    function Harness() {
      const hostRef = useRef<HTMLDivElement | null>(null)
      const liveGeometry = useMemo(
        () =>
          createGridGeometrySnapshotFromAxes({
            columns: createGridAxisWorldIndex({ axisLength: 300, defaultSize: gridMetrics.columnWidth }),
            dpr: 1,
            gridMetrics,
            hostHeight: 420,
            hostWidth: 640,
            rows: createGridAxisWorldIndex({ axisLength: 300, defaultSize: gridMetrics.rowHeight }),
            scrollLeft: 0,
            scrollTop: 0,
            sheetName: 'Sheet1',
            updatedAt: 100,
          }),
        [],
      )
      latestResolvers = useWorkbookGridPointerResolvers({
        getGeometrySnapshot: () => liveGeometry,
        gridSelection: createGridSelection(0, 0),
        hostRef,
        selectedCell: { col: 0, row: 0 },
      })

      return (
        <>
          <div
            ref={(node) => {
              hostRef.current = node
              if (node) {
                node.getBoundingClientRect = () =>
                  ({
                    bottom: 420,
                    height: 420,
                    left: 10,
                    right: 650,
                    top: 10,
                    width: 640,
                    x: 10,
                    y: 10,
                    toJSON: () => ({}),
                  }) as DOMRect
              }
            }}
          />
        </>
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
    })

    expect(
      latestResolvers?.resolveHeaderSelectionAtPointer(10 + gridMetrics.rowMarkerWidth + gridMetrics.columnWidth * 2 + 1, 10 + 12),
    ).toEqual({ kind: 'column', index: 2 })

    await act(async () => {
      root.unmount()
    })
  })
})
