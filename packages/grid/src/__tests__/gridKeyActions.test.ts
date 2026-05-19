import { describe, expect, test } from 'vitest'
import { resolveGridKeyAction } from '../gridKeyActions.js'

describe('gridKeyActions', () => {
  test('appends printable characters during edit mode when the editor input is not focused', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'x', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: 'abc',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: 'edit-append', value: 'abcx' })
  })

  test('ignores printable characters during edit mode when the editor input is focused', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'x', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: 'abc',
        editorInputFocused: true,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: 'none' })
  })

  test('commits and cancels edit mode keys before the overlay input takes focus', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'Enter', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: '123',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: 'commit-edit', movement: [0, 1] })

    expect(
      resolveGridKeyAction({
        event: { key: 'Tab', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: true,
        editorValue: '123',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: 'commit-edit', movement: [-1, 0] })

    expect(
      resolveGridKeyAction({
        event: { key: 'Escape', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: '123',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: 'cancel-edit' })
  })

  test('returns begin-edit and movement actions for core keys', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'F2', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: '12',
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: 'begin-edit', selectionBehavior: 'caret-end', pendingTypeSeed: null })

    expect(
      resolveGridKeyAction({
        event: { key: 'Enter', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: 'move-selection', cell: [2, 5] })

    expect(
      resolveGridKeyAction({
        event: { key: 'ArrowRight', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [1, 4],
      }),
    ).toEqual({ kind: 'extend-selection', anchor: [1, 4], target: [3, 4] })
  })

  test('cycles Enter and Tab through the active range without collapsing it', () => {
    const range = { x: 1, y: 1, width: 2, height: 2 }

    expect(
      resolveGridKeyAction({
        event: { key: 'Tab', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: range,
      }),
    ).toEqual({ kind: 'move-selection-in-range', cell: [2, 1], range })

    expect(
      resolveGridKeyAction({
        event: { key: 'Tab', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: range,
      }),
    ).toEqual({ kind: 'move-selection-in-range', cell: [2, 2], range })

    expect(
      resolveGridKeyAction({
        event: { key: 'Enter', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: range,
      }),
    ).toEqual({ kind: 'move-selection-in-range', cell: [1, 2], range })

    expect(
      resolveGridKeyAction({
        event: { key: 'Enter', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [2, 2],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: range,
      }),
    ).toEqual({ kind: 'move-selection-in-range', cell: [1, 1], range })
  })

  test('supports sheet-style navigation keys and selection shortcuts', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'Home', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [5, 4],
        currentSelectionCell: [5, 4],
        currentRangeAnchor: [5, 4],
      }),
    ).toEqual({ kind: 'move-selection', cell: [0, 4] })

    expect(
      resolveGridKeyAction({
        event: { key: 'End', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [5, 4],
        currentSelectionCell: [5, 4],
        currentRangeAnchor: [2, 1],
      }),
    ).toEqual({ kind: 'extend-selection', anchor: [2, 1], target: [16383, 1048575] })

    expect(
      resolveGridKeyAction({
        event: { key: 'PageDown', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: 'move-selection', cell: [2, 24] })

    expect(
      resolveGridKeyAction({
        event: { key: ' ', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: 'select-row', col: 3, row: 9 })

    expect(
      resolveGridKeyAction({
        event: { key: ' ', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: 'select-column', col: 3, row: 9 })

    expect(
      resolveGridKeyAction({
        event: {
          key: ' ',
          ctrlKey: true,
          metaKey: false,
          altKey: false,
          shiftKey: true,
        },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: 'select-all' })

    expect(
      resolveGridKeyAction({
        event: { key: 'a', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: 'select-all' })
  })

  test('uses data-aware navigation resolvers for spreadsheet table-stakes shortcuts', () => {
    const navigation = {
      resolveCurrentRegion: () => ({ x: 1, y: 2, width: 3, height: 4 }),
      resolveDataEdge: (): [number, number] => [2, 9],
    }

    expect(
      resolveGridKeyAction({
        event: { key: 'ArrowDown', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        navigation,
      }),
    ).toEqual({ kind: 'move-selection', cell: [2, 9] })

    expect(
      resolveGridKeyAction({
        event: { key: 'ArrowDown', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [1, 4],
        navigation,
      }),
    ).toEqual({ kind: 'extend-selection', anchor: [1, 4], target: [2, 9] })

    expect(
      resolveGridKeyAction({
        event: { key: '*', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        navigation,
      }),
    ).toEqual({ kind: 'select-range', cell: [2, 4], range: { x: 1, y: 2, width: 3, height: 4 } })

    expect(
      resolveGridKeyAction({
        event: { key: 'a', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        navigation,
      }),
    ).toEqual({ kind: 'select-range', cell: [2, 4], range: { x: 1, y: 2, width: 3, height: 4 } })

    expect(
      resolveGridKeyAction({
        event: { key: 'a', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        currentSelectionRange: { x: 1, y: 2, width: 3, height: 4 },
        navigation,
      }),
    ).toEqual({ kind: 'select-all' })
  })

  test('handles select-current-region without falling through to printable edit when no region exists', () => {
    expect(
      resolveGridKeyAction({
        event: { key: '*', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        navigation: {
          resolveCurrentRegion: () => null,
          resolveDataEdge: () => null,
        },
      }),
    ).toEqual({ kind: 'handled' })
  })

  test('routes Google Sheets structural delete shortcut for selected rows and columns', () => {
    expect(
      resolveGridKeyAction({
        event: { key: '-', ctrlKey: true, metaKey: false, altKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 3],
        currentSelectionCell: [1, 3],
        currentRangeAnchor: [1, 3],
        selectedRowRanges: [{ start: 3, count: 2 }],
      }),
    ).toEqual({ kind: 'delete-selected-rows', ranges: [{ start: 3, count: 2 }] })

    expect(
      resolveGridKeyAction({
        event: { key: '-', ctrlKey: false, metaKey: true, altKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 1],
        currentSelectionCell: [2, 1],
        currentRangeAnchor: [2, 1],
        selectedColumnRanges: [{ start: 2, count: 3 }],
      }),
    ).toEqual({ kind: 'delete-selected-columns', ranges: [{ start: 2, count: 3 }] })

    expect(
      resolveGridKeyAction({
        event: { key: '-', ctrlKey: true, metaKey: false, altKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 1],
        currentSelectionCell: [2, 1],
        currentRangeAnchor: [2, 1],
        selectedRowRanges: [{ start: 0, count: 10 }],
        selectedColumnRanges: [{ start: 0, count: 10 }],
      }),
    ).toEqual({ kind: 'handled' })
  })

  test('moves PageUp and PageDown by the caller-provided visible viewport height', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'PageDown', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        pageJumpRows: 31,
      }),
    ).toEqual({ kind: 'move-selection', cell: [2, 35] })
  })

  test('extends an existing keyboard range from the active edge', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'ArrowRight', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        currentSelectionRange: { x: 2, y: 4, width: 2, height: 1 },
      }),
    ).toEqual({ kind: 'extend-selection', anchor: [2, 4], target: [4, 4] })

    expect(
      resolveGridKeyAction({
        event: { key: 'ArrowDown', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        currentSelectionRange: { x: 2, y: 4, width: 3, height: 2 },
      }),
    ).toEqual({ kind: 'extend-selection', anchor: [2, 4], target: [4, 1048575] })
  })

  test('returns clipboard and typed-entry actions', () => {
    expect(
      resolveGridKeyAction({
        event: { key: 'c', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [3, 2],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: 'clipboard-copy' })

    expect(
      resolveGridKeyAction({
        event: { key: 'v', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [3, 2],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: 'clipboard-paste', target: [3, 2], valuesOnly: false })

    expect(
      resolveGridKeyAction({
        event: { key: 'v', ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [3, 2],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: 'clipboard-paste', target: [3, 2], valuesOnly: true })

    expect(
      resolveGridKeyAction({
        event: { key: 'Enter', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [2, 2],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: { x: 1, y: 1, width: 3, height: 4 },
      }),
    ).toEqual({
      kind: 'fill-range',
      source: { x: 2, y: 2, width: 1, height: 1 },
      target: { x: 1, y: 1, width: 3, height: 4 },
    })

    expect(
      resolveGridKeyAction({
        event: { key: 'd', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: { x: 1, y: 1, width: 3, height: 4 },
      }),
    ).toEqual({
      kind: 'fill-range',
      source: { x: 1, y: 1, width: 3, height: 1 },
      target: { x: 1, y: 2, width: 3, height: 3 },
    })

    expect(
      resolveGridKeyAction({
        event: { key: 'r', ctrlKey: false, metaKey: true, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: { x: 1, y: 1, width: 3, height: 4 },
      }),
    ).toEqual({
      kind: 'fill-range',
      source: { x: 1, y: 1, width: 1, height: 4 },
      target: { x: 2, y: 1, width: 2, height: 4 },
    })

    expect(
      resolveGridKeyAction({
        event: { key: 'd', ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
        currentSelectionRange: { x: 1, y: 1, width: 3, height: 1 },
      }),
    ).toEqual({ kind: 'handled' })

    expect(
      resolveGridKeyAction({
        event: { key: '7', ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: '',
        editorInputFocused: false,
        pendingTypeSeed: '1',
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({
      kind: 'begin-edit',
      seed: '17',
      selectionBehavior: 'caret-end',
      pendingTypeSeed: '17',
    })
  })

  test('does not clear cells for modified delete key combinations', () => {
    for (const event of [
      { key: 'Delete', ctrlKey: true, metaKey: false, altKey: false },
      { key: 'Delete', ctrlKey: false, metaKey: true, altKey: false },
      { key: 'Delete', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'Delete', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
      { key: 'Backspace', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'Backspace', ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
    ] as const) {
      expect(
        resolveGridKeyAction({
          event,
          isEditingCell: false,
          editorValue: '',
          editorInputFocused: false,
          pendingTypeSeed: null,
          selectedCell: [1, 1],
          currentSelectionCell: [1, 1],
          currentRangeAnchor: [1, 1],
        }),
      ).toEqual({ kind: 'none' })
    }
  })

  test('routes primary-modified Backspace to scroll the active cell into view without clearing it', () => {
    for (const event of [
      { key: 'Backspace', ctrlKey: true, metaKey: false, altKey: false },
      { key: 'Backspace', ctrlKey: false, metaKey: true, altKey: false },
    ] as const) {
      expect(
        resolveGridKeyAction({
          event,
          isEditingCell: false,
          editorValue: '',
          editorInputFocused: false,
          pendingTypeSeed: null,
          selectedCell: [1, 1],
          currentSelectionCell: [1, 1],
          currentRangeAnchor: [1, 1],
        }),
      ).toEqual({ kind: 'scroll-active-cell' })
    }
  })

  test('does not claim unadvertised modified navigation and edit shortcuts', () => {
    for (const event of [
      { key: 'ArrowLeft', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'ArrowRight', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'Home', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'End', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'PageDown', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'Tab', ctrlKey: false, metaKey: true, altKey: false },
      { key: 'F2', ctrlKey: false, metaKey: false, altKey: true },
    ] as const) {
      expect(
        resolveGridKeyAction({
          event,
          isEditingCell: false,
          editorValue: '',
          editorInputFocused: false,
          pendingTypeSeed: null,
          selectedCell: [1, 1],
          currentSelectionCell: [1, 1],
          currentRangeAnchor: [1, 1],
        }),
      ).toEqual({ kind: 'none' })
    }
  })

  test('does not commit edit mode for modified Enter or Tab before the editor input owns focus', () => {
    for (const event of [
      { key: 'Enter', ctrlKey: false, metaKey: false, altKey: true },
      { key: 'Enter', ctrlKey: true, metaKey: false, altKey: false },
      { key: 'Tab', ctrlKey: false, metaKey: true, altKey: false },
      { key: 'Escape', ctrlKey: false, metaKey: false, altKey: true },
    ] as const) {
      expect(
        resolveGridKeyAction({
          event,
          isEditingCell: true,
          editorValue: 'draft',
          editorInputFocused: false,
          pendingTypeSeed: null,
          selectedCell: [1, 1],
          currentSelectionCell: [1, 1],
          currentRangeAnchor: [1, 1],
        }),
      ).toEqual({ kind: 'none' })
    }
  })
})
