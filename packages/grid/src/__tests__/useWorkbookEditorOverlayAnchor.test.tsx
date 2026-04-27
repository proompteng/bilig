// @vitest-environment jsdom
import { act, useCallback, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { Rectangle } from '../gridTypes.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'
import { useWorkbookEditorOverlayAnchor, type WorkbookEditorOverlayAnchorState } from '../useWorkbookEditorOverlayAnchor.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'

function createSnapshot(styleId?: string): CellSnapshot {
  return {
    address: 'B2',
    flags: 0,
    sheetName: 'Sheet1',
    styleId,
    value: { tag: ValueTag.String, value: 'draft' },
    version: 1,
  }
}

function createHostRect(): DOMRect {
  return {
    bottom: 500,
    height: 300,
    left: 100,
    right: 600,
    toJSON: () => ({}),
    top: 200,
    width: 500,
    x: 100,
    y: 200,
  }
}

describe('useWorkbookEditorOverlayAnchor', () => {
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalCancelAnimationFrame = window.cancelAnimationFrame

  beforeEach(() => {
    window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
      const handle = window.setTimeout(() => callback(performance.now()), 0)
      return handle
    }) as typeof window.requestAnimationFrame
    window.cancelAnimationFrame = ((handle: number) => {
      window.clearTimeout(handle)
    }) as typeof window.cancelAnimationFrame
  })

  afterEach(() => {
    window.requestAnimationFrame = originalRequestAnimationFrame
    window.cancelAnimationFrame = originalCancelAnimationFrame
    document.body.innerHTML = ''
  })

  it('anchors the mounted editor to the selected cell and updates scroll movement without a React commit', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const cameraStore = new GridCameraStore()
    const scrollStore = new WorkbookGridScrollStore()
    const engine: GridEngineLike = {
      workbook: {
        getSheet: () => undefined,
      },
      getCell: () => createSnapshot(),
      getCellStyle: () => ({
        fill: { backgroundColor: '#1f2937' },
        font: { color: '#f8fafc' },
        id: 'style-1',
      }),
      subscribeCells: () => () => undefined,
    }
    let localBounds: Rectangle = { height: 24, width: 104, x: 20, y: 30 }
    let latestState: WorkbookEditorOverlayAnchorState | null = null
    let committedStyleBeforeScroll: WorkbookEditorOverlayAnchorState['overlayStyle']

    function Harness() {
      const [hostElement, setHostElement] = useState<HTMLDivElement | null>(null)
      const handleHostRef = useCallback((node: HTMLDivElement | null) => {
        if (node) {
          node.getBoundingClientRect = createHostRect
        }
        setHostElement(node)
      }, [])
      const state = useWorkbookEditorOverlayAnchor({
        editorValue: '123',
        engine,
        getCellLocalBounds: () => localBounds,
        gridCameraStore: cameraStore,
        hostElement,
        isEditingCell: true,
        scrollTransformStore: scrollStore,
        selectedCellSnapshot: createSnapshot('style-1'),
        selectedCol: 1,
        selectedRow: 1,
      })
      latestState = state
      return (
        <div ref={handleHostRef}>
          <div data-testid="cell-editor-overlay" style={state.overlayStyle} />
        </div>
      )
    }

    const rootHost = document.createElement('div')
    document.body.appendChild(rootHost)
    const root = createRoot(rootHost)

    await act(async () => {
      root.render(<Harness />)
      await new Promise((resolve) => window.setTimeout(resolve, 0))
    })

    const overlay = document.querySelector<HTMLElement>('[data-testid="cell-editor-overlay"]')
    expect(overlay?.style.left).toBe('120px')
    expect(overlay?.style.top).toBe('230px')
    expect(overlay?.style.width).toBe('104px')
    expect(overlay?.style.height).toBe('24px')
    expect(latestState?.editorTextAlign).toBe('right')
    expect(latestState?.editorPresentation.backgroundColor).toBe('#1f2937')
    committedStyleBeforeScroll = latestState?.overlayStyle

    localBounds = { height: 24, width: 104, x: 37, y: 44 }
    scrollStore.setSnapshot({ renderTx: 17, renderTy: 14, scrollLeft: 17, scrollTop: 14, tx: 17, ty: 14 })

    expect(overlay?.style.left).toBe('137px')
    expect(overlay?.style.top).toBe('244px')
    expect(latestState?.overlayStyle).toBe(committedStyleBeforeScroll)

    await act(async () => {
      root.unmount()
    })
  })
})
