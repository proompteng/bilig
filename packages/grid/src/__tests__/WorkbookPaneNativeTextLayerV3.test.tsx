// @vitest-environment jsdom
import { describe, expect, test } from 'vitest'
import {
  resolveNativeTextRunFontStyleV3,
  resolveNativeTextRunInnerStyleV3,
  resolveNativeTextRunOuterStyleV3,
  resolveNativeTextRunVisibleClipV3,
} from '../renderer-v3/WorkbookPaneNativeTextLayerV3.js'
import type { TextQuadRun } from '../renderer-v3/line-text-quad-buffer.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'

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

function createPane(): WorkbookRenderTilePaneState {
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
      lastBatchId: 1,
      lastCameraSeq: 1,
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 1,
      textMetrics: new Float32Array(),
      textRuns: [createRun()],
      tileId: 1,
      version: { axisX: 1, axisY: 1, freeze: 0, styles: 1, text: 1, values: 1 },
    },
    viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
  }
}

describe('WorkbookPaneNativeTextLayerV3', () => {
  test('snaps text layer placement to device pixels while preserving run clipping', () => {
    expect(
      resolveNativeTextRunOuterStyleV3({
        dpr: 2,
        pane: createPane(),
        run: createRun({ clipX: 8.2, clipY: 4.3 }),
        scrollSnapshot: { renderTx: 1.1, renderTy: 2.2, tx: 1.1, ty: 2.2 },
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
      fontSize: 14.667,
      fontStyle: 'normal',
      fontVariantNumeric: 'tabular-nums',
      fontWeight: 400,
      height: 17.5,
      letterSpacing: 0,
      lineHeight: '17.5px',
      textDecorationLine: 'underline',
      textAlign: 'right',
      top: -1.5,
      whiteSpace: 'pre',
    })
    expect(style).not.toHaveProperty('MozOsxFontSmoothing')
    expect(style).not.toHaveProperty('textRendering')
    expect(style).not.toHaveProperty('WebkitFontSmoothing')
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
})
