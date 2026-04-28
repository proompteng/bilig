// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { Rectangle } from '../gridTypes.js'
import { GridEditorAnchorRuntime } from '../runtime/gridEditorAnchorRuntime.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'

function createSnapshot(styleId?: string): CellSnapshot {
  return {
    address: 'B2',
    flags: 0,
    input: '',
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

function createEngine(): GridEngineLike {
  return {
    getCell: () => createSnapshot(),
    getCellStyle: () => ({
      fill: { backgroundColor: '#1f2937' },
      font: { color: '#f8fafc' },
      id: 'style-1',
    }),
    subscribeCells: () => () => undefined,
    workbook: {
      getSheet: () => undefined,
    },
  }
}

describe('GridEditorAnchorRuntime', () => {
  it('resolves and applies editor overlay bounds without requiring React state', () => {
    const runtime = new GridEditorAnchorRuntime()
    const cameraStore = new GridCameraStore()
    const host = document.createElement('div')
    host.getBoundingClientRect = createHostRect
    document.body.innerHTML = '<div data-testid="cell-editor-overlay"></div>'
    const localBounds: Rectangle = { height: 24, width: 104, x: 20, y: 30 }

    const bounds = runtime.refreshOverlayBounds({
      col: 1,
      getCellLocalBounds: () => localBounds,
      gridCameraStore: cameraStore,
      hostElement: host,
      row: 1,
    })

    const overlay = document.querySelector<HTMLElement>('[data-testid="cell-editor-overlay"]')
    expect(bounds).toEqual({ height: 24, width: 104, x: 120, y: 230 })
    expect(overlay?.style.left).toBe('120px')
    expect(overlay?.style.top).toBe('230px')
    expect(overlay?.style.width).toBe('104px')
    expect(overlay?.style.height).toBe('24px')
  })

  it('keeps committed bounds stable when the resolved bounds have not changed', () => {
    const runtime = new GridEditorAnchorRuntime()
    const current: Rectangle = { height: 24, width: 104, x: 120, y: 230 }

    expect(runtime.resolveCommittedBounds(current, { height: 24, width: 104, x: 120, y: 230 })).toBe(current)
    expect(runtime.resolveCommittedBounds(current, { height: 24, width: 104, x: 121, y: 230 })).toEqual({
      height: 24,
      width: 104,
      x: 121,
      y: 230,
    })
  })

  it('resolves editor presentation and alignment outside the React hook', () => {
    const runtime = new GridEditorAnchorRuntime()
    const presentation = runtime.resolvePresentation({
      engine: createEngine(),
      selectedCellSnapshot: createSnapshot('style-1'),
    })

    expect(presentation.backgroundColor).toBe('#1f2937')
    expect(presentation.color).toBe('#f8fafc')
    expect(runtime.resolveTextAlign('123')).toBe('right')
    expect(runtime.resolveTextAlign('draft')).toBe('left')
    expect(runtime.resolveOverlayStyle(true, { height: 24, width: 104, x: 120, y: 230 })).toMatchObject({
      height: 24,
      left: 120,
      position: 'fixed',
      top: 230,
      width: 104,
    })
  })
})
