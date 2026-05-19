// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, test } from 'vitest'
import {
  WorkbookPaneNativeTextLayerV3,
  resolveNativeTextLayerDrawScrollSnapshotV3,
  resolveNativeTextRunFontStyleV3,
  resolveNativeTextRunInnerStyleV3,
  resolveNativeTextRunOuterStyleV3,
  resolveNativeTextRunVisibleClipV3,
} from '../renderer-v3/WorkbookPaneNativeTextLayerV3.js'
import type { TextQuadRun } from '../renderer-v3/line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'

function createRun(overrides: Partial<TextQuadRun> = {}): TextQuadRun {
  return {
    align: 'left',
    clipHeight: 18,
    clipWidth: 90,
    clipX: 8,
    clipY: 4,
    color: '#1f2933',
    font: '400 14.667px Arial, sans-serif',
    fontSize: 14.667,
    height: 22,
    strike: false,
    text: 'Annual software',
    underline: false,
    width: 104,
    wrap: false,
    x: 0,
    y: 0,
    ...overrides,
  }
}

function createPane(
  run: TextQuadRun | readonly TextQuadRun[] = createRun(),
  overrides: {
    readonly lastBatchId?: number | undefined
    readonly lastCameraSeq?: number | undefined
    readonly version?: Partial<WorkbookRenderTilePaneState['tile']['version']> | undefined
  } = {},
): WorkbookRenderTilePaneState {
  const textRuns = Array.isArray(run) ? run : [run]
  return {
    contentOffset: { x: 0, y: 0 },
    frame: { height: 240, width: 320, x: 46, y: 24 },
    generation: 1,
    paneId: 'body',
    scrollAxes: { x: true, y: true },
    surfaceSize: { height: 240, width: 320 },
    tile: {
      bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      coord: { colTile: 0, dprBucket: 2, paneKind: 'body', rowTile: 0, sheetId: 1, sheetOrdinal: 1 },
      lastBatchId: overrides.lastBatchId ?? 1,
      lastCameraSeq: overrides.lastCameraSeq ?? 1,
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: textRuns.length,
      textMetrics: new Float32Array(),
      textRuns,
      tileId: 1,
      version: { axisX: 1, axisY: 1, freeze: 0, styles: 1, text: 1, values: 1, ...overrides.version },
    },
    viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
  }
}

describe('WorkbookPaneNativeTextLayerV3', () => {
  test('snaps text clip bounds while preserving fractional glyph origin inside the clip', () => {
    const visibleClip = resolveNativeTextRunVisibleClipV3({
      dpr: 2,
      pane: createPane(),
      run: createRun({ clipX: 8.2, clipY: 4.3 }),
      scrollSnapshot: { renderTx: 1.1, renderTy: 2.2, tx: 1.1, ty: 2.2 },
    })

    expect(visibleClip).toMatchObject({
      outerHeight: 18,
      outerLeft: 53,
      outerTop: 26,
      outerWidth: 90,
    })
    expect(visibleClip?.innerLeft).toBeCloseTo(-8.1)
    expect(visibleClip?.innerTop).toBeCloseTo(-4.2)
    expect(
      resolveNativeTextRunOuterStyleV3({
        dpr: 2,
        pane: createPane(),
        run: createRun({ clipX: 8.2, clipY: 4.3 }),
        scrollSnapshot: { renderTx: 1.1, renderTy: 2.2, tx: 1.1, ty: 2.2 },
        visibleClip,
      }),
    ).toMatchObject({
      height: 18,
      left: 53,
      overflow: 'hidden',
      top: 26,
      width: 90,
    })
  })

  test('uses native browser font rendering with spreadsheet numeric alignment', () => {
    const style = resolveNativeTextRunInnerStyleV3({ dpr: 2, run: createRun({ align: 'right', underline: true }) })
    expect(style).toMatchObject({
      color: '#1f2933',
      display: 'block',
      fontFamily: 'Arial, sans-serif',
      fontFeatureSettings: 'normal',
      fontOpticalSizing: 'auto',
      fontSize: 14.667,
      fontSynthesis: 'none',
      fontStyle: 'normal',
      fontVariantNumeric: 'tabular-nums',
      fontWeight: 400,
      height: 17.5,
      letterSpacing: 0,
      lineHeight: '17.5px',
      MozOsxFontSmoothing: 'grayscale',
      textDecorationLine: 'underline',
      textAlign: 'right',
      textRendering: 'optimizeLegibility',
      top: -1.5,
      WebkitFontSmoothing: 'antialiased',
      whiteSpace: 'pre',
    })
  })

  test('keeps wrapped text top-aligned while non-wrapped text uses a snapped line box', () => {
    expect(resolveNativeTextRunInnerStyleV3({ dpr: 2, run: createRun({ wrap: true }) })).toMatchObject({
      display: 'block',
      height: 22,
      lineHeight: '17.5px',
      top: -4,
      whiteSpace: 'pre-wrap',
    })
  })

  test('parses styled spreadsheet font strings into stable CSS longhands', () => {
    expect(
      resolveNativeTextRunFontStyleV3(createRun({ font: 'italic 700 18.667px Aptos, Arial, sans-serif', fontSize: undefined })),
    ).toEqual({
      fontFamily: 'Aptos, Arial, sans-serif',
      fontSize: 18.667,
      fontStyle: 'italic',
      fontWeight: 700,
    })
  })

  test('clips long spill text to the visible pane frame', () => {
    const pane = createPane()
    const run = createRun({
      clipWidth: 13_000,
      clipX: 152,
      width: 13_000,
      x: 152,
    })
    const visibleClip = resolveNativeTextRunVisibleClipV3({
      dpr: 1,
      pane,
      run,
      scrollSnapshot: { tx: 0, ty: 0 },
    })

    expect(visibleClip).toMatchObject({
      innerLeft: 0,
      innerWidth: 168,
      outerLeft: 198,
      outerWidth: 168,
    })
    expect(resolveNativeTextRunOuterStyleV3({ dpr: 1, pane, run, scrollSnapshot: { tx: 0, ty: 0 }, visibleClip })).toMatchObject({
      left: 198,
      width: 168,
    })
    expect(resolveNativeTextRunInnerStyleV3({ dpr: 1, run, visibleClip })).toMatchObject({
      left: 0,
      width: 168,
    })
  })

  test('drops spill text runs that do not intersect the pane frame', () => {
    expect(
      resolveNativeTextRunVisibleClipV3({
        dpr: 1,
        pane: createPane(),
        run: createRun({ clipWidth: 400, clipX: 400, width: 400, x: 400 }),
        scrollSnapshot: { tx: 0, ty: 0 },
      }),
    ).toBeNull()
  })

  test('uses the presented renderer scroll frame instead of racing ahead on live scroll', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const scrollStore = new WorkbookGridScrollStore()
    const pane = createPane(createRun({ clipX: 40, col: 1, row: 2, text: 'stable-frame', x: 40 }))
    const presentedAtOrigin = { renderTx: 0, renderTy: 0, scrollLeft: 0, scrollTop: 0, tx: 0, ty: 0 }
    const presentedAfterScroll = { renderTx: 20, renderTy: 8, scrollLeft: 20, scrollTop: 8, tx: 20, ty: 8 }
    const readRunLeft = () => host.querySelector<HTMLElement>('[data-native-text-run]')?.style.left ?? null
    const readLayer = () => host.querySelector<HTMLElement>('[data-testid="grid-native-text-layer"]')

    try {
      scrollStore.setSnapshot(presentedAfterScroll)

      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            presentedScrollSnapshot={presentedAtOrigin}
            scrollTransformStore={scrollStore}
            tilePanes={[pane]}
          />,
        )
      })

      expect(readRunLeft()).toBe('86px')
      expect(readLayer()?.getAttribute('data-v3-native-text-render-tx')).toBe('0')
      expect(
        resolveNativeTextLayerDrawScrollSnapshotV3({
          geometry: null,
          liveScrollSnapshot: presentedAfterScroll,
          panes: [pane],
          presentedScrollSnapshot: presentedAtOrigin,
        }),
      ).toBe(presentedAtOrigin)

      await act(async () => {
        scrollStore.setSnapshot({ renderTx: 48, renderTy: 16, scrollLeft: 48, scrollTop: 16, tx: 48, ty: 16 })
      })

      expect(readRunLeft()).toBe('86px')
      expect(readLayer()?.getAttribute('data-v3-native-text-render-tx')).toBe('0')

      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            presentedScrollSnapshot={presentedAfterScroll}
            scrollTransformStore={scrollStore}
            tilePanes={[pane]}
          />,
        )
      })

      expect(readRunLeft()).toBe('66px')
      expect(readLayer()?.getAttribute('data-v3-native-text-render-tx')).toBe('20')
    } finally {
      await act(async () => {
        root.unmount()
      })
      host.remove()
    }
  })

  test('rerenders snapped native text geometry when device pixel ratio changes', async () => {
    const originalDevicePixelRatio = Object.getOwnPropertyDescriptor(window, 'devicePixelRatio')
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const readRunLeft = () => {
      const run = host.querySelector<HTMLElement>('[data-native-text-run]')
      return run?.style.left ?? null
    }

    try {
      Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 1 })

      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            tilePanes={[createPane(createRun({ clipX: 8.3 }))]}
          />,
        )
      })

      expect(readRunLeft()).toBe('54px')

      Object.defineProperty(window, 'devicePixelRatio', { configurable: true, value: 2 })
      await act(async () => {
        window.dispatchEvent(new Event('resize'))
      })

      expect(readRunLeft()).toBe('54.5px')
    } finally {
      await act(async () => {
        root.unmount()
      })
      host.remove()
      if (originalDevicePixelRatio) {
        Object.defineProperty(window, 'devicePixelRatio', originalDevicePixelRatio)
      }
    }
  })

  test('suppresses the active editor cell from the native text layer while keeping neighboring text visible', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            suppressedTextCell={{ col: 1, row: 2 }}
            tilePanes={[
              createPane([
                createRun({ col: 1, row: 2, text: 'editing-cell' }),
                createRun({ col: 2, row: 2, text: 'neighbor-cell', x: 104, clipX: 104 }),
              ]),
            ]}
          />,
        )
      })

      const textLayer = host.querySelector<HTMLElement>('[data-testid="grid-native-text-layer"]')
      expect(textLayer?.getAttribute('data-v3-native-text-run-count')).toBe('1')
      expect(host.querySelector('[data-native-text-run-row="2"][data-native-text-run-col="1"]')).toBeNull()
      expect(host.querySelector('[data-native-text-run-row="2"][data-native-text-run-col="2"]')?.textContent).toBe('neighbor-cell')
    } finally {
      await act(async () => {
        root.unmount()
      })
      host.remove()
    }
  })

  test('keeps text DOM nodes stable across tile revision churn', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const stableRun = createRun({ col: 1, row: 2, text: 'stable-cell' })

    try {
      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            tilePanes={[createPane(stableRun)]}
          />,
        )
      })

      const firstRun = host.querySelector<HTMLElement>('[data-native-text-run-row="2"][data-native-text-run-col="1"]')
      expect(firstRun?.textContent).toBe('stable-cell')

      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            tilePanes={[
              createPane(stableRun, {
                lastBatchId: 42,
                lastCameraSeq: 11,
                version: { styles: 7, text: 8, values: 9 },
              }),
            ]}
          />,
        )
      })

      expect(host.querySelector('[data-native-text-run-row="2"][data-native-text-run-col="1"]')).toBe(firstRun)
    } finally {
      await act(async () => {
        root.unmount()
      })
      host.remove()
    }
  })

  test('updates edited text in place instead of remounting the cell run', async () => {
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            tilePanes={[createPane(createRun({ col: 1, row: 2, text: 'before-edit' }))]}
          />,
        )
      })

      const firstRun = host.querySelector<HTMLElement>('[data-native-text-run-row="2"][data-native-text-run-col="1"]')
      expect(firstRun?.textContent).toBe('before-edit')

      await act(async () => {
        root.render(
          <WorkbookPaneNativeTextLayerV3
            active
            cameraStore={null}
            geometry={null}
            headerPanes={[]}
            scrollTransformStore={null}
            tilePanes={[
              createPane(createRun({ col: 1, row: 2, text: 'after-edit' }), {
                lastBatchId: 2,
                version: { text: 2, values: 2 },
              }),
            ]}
          />,
        )
      })

      const updatedRun = host.querySelector<HTMLElement>('[data-native-text-run-row="2"][data-native-text-run-col="1"]')
      expect(updatedRun).toBe(firstRun)
      expect(updatedRun?.textContent).toBe('after-edit')
    } finally {
      await act(async () => {
        root.unmount()
      })
      host.remove()
    }
  })
})
