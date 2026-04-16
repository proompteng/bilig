import { describe, expect, test } from 'vitest'
import { resolveBodyDoubleClickIntent, resolveHeaderClickIntent, shouldSkipGridSelectionChange } from '../gridEventPolicy.js'

describe('gridEventPolicy', () => {
  test('resolves body double click into ignore, edit, or autofit intent', () => {
    expect(
      resolveBodyDoubleClickIntent({
        resizeTarget: null,
        bodyCell: [2, 4],
        lastBodyClickCell: [2, 3],
      }),
    ).toEqual({ kind: 'ignore' })

    expect(
      resolveBodyDoubleClickIntent({
        resizeTarget: 3,
        bodyCell: [2, 4],
        lastBodyClickCell: [2, 4],
      }),
    ).toEqual({ kind: 'autofit-column', columnIndex: 3 })

    expect(
      resolveBodyDoubleClickIntent({
        resizeTarget: null,
        bodyCell: [2, 4],
        lastBodyClickCell: [2, 4],
      }),
    ).toEqual({ kind: 'edit-cell', cell: [2, 4] })

    expect(
      resolveBodyDoubleClickIntent({
        resizeTarget: null,
        bodyCell: [2, 4],
        lastBodyClickCell: [2, 3],
      }),
    ).toEqual({ kind: 'ignore' })
  })

  test('resolves header click into ignore, autofit, or selection intent', () => {
    expect(
      resolveHeaderClickIntent({
        isEdge: true,
        isDoubleClick: true,
        columnResizeActive: false,
        columnIndex: 4,
        selectedRow: 7,
      }),
    ).toEqual({ kind: 'autofit-column', columnIndex: 4 })

    expect(
      resolveHeaderClickIntent({
        isEdge: false,
        isDoubleClick: false,
        columnResizeActive: true,
        columnIndex: 4,
        selectedRow: 7,
      }),
    ).toEqual({ kind: 'ignore' })

    expect(
      resolveHeaderClickIntent({
        isEdge: false,
        isDoubleClick: false,
        columnResizeActive: false,
        columnIndex: 4,
        selectedRow: 7,
      }),
    ).toEqual({ kind: 'select-column', addr: 'E8', columnIndex: 4, selectedRow: 7 })
  })

  test('marks grid selection changes that should be skipped', () => {
    expect(
      shouldSkipGridSelectionChange({
        columnResizeActive: true,
        postDragSelectionExpiry: 0,
        now: 100,
        ignoreNextPointerSelection: false,
        hasDragViewport: false,
      }),
    ).toEqual({
      skip: true,
      consumeIgnoreNextPointerSelection: false,
      clearPostDragSelectionExpiry: false,
    })

    expect(
      shouldSkipGridSelectionChange({
        columnResizeActive: false,
        postDragSelectionExpiry: 120,
        now: 100,
        ignoreNextPointerSelection: false,
        hasDragViewport: false,
      }),
    ).toEqual({
      skip: true,
      consumeIgnoreNextPointerSelection: false,
      clearPostDragSelectionExpiry: true,
    })

    expect(
      shouldSkipGridSelectionChange({
        columnResizeActive: false,
        postDragSelectionExpiry: 0,
        now: 100,
        ignoreNextPointerSelection: true,
        hasDragViewport: false,
      }),
    ).toEqual({
      skip: true,
      consumeIgnoreNextPointerSelection: true,
      clearPostDragSelectionExpiry: false,
    })

    expect(
      shouldSkipGridSelectionChange({
        columnResizeActive: false,
        postDragSelectionExpiry: 0,
        now: 100,
        ignoreNextPointerSelection: false,
        hasDragViewport: false,
      }),
    ).toEqual({
      skip: false,
      consumeIgnoreNextPointerSelection: false,
      clearPostDragSelectionExpiry: false,
    })
  })
})
