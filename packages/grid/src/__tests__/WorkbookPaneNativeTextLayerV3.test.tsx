// @vitest-environment jsdom
import { describe, expect, test } from 'vitest'
import { resolveNativeTextRunInnerStyleV3, resolveNativeTextRunOuterStyleV3 } from '../renderer-v3/WorkbookPaneNativeTextLayerV3.js'
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

  test('uses native browser font rendering styles for visible workbook text', () => {
    expect(resolveNativeTextRunInnerStyleV3({ dpr: 2, run: createRun({ align: 'right', underline: true }) })).toMatchObject({
      alignItems: 'center',
      color: '#1f2933',
      font: '400 14.667px Arial, sans-serif',
      justifyContent: 'flex-end',
      textDecorationLine: 'underline',
      textRendering: 'auto',
      whiteSpace: 'pre',
      WebkitFontSmoothing: 'auto',
    })
  })
})
