// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createGridSelection, createRangeSelection } from '../gridSelection.js'
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  handleGridKey,
  handleGridPasteCapture,
  isGridKeyboardEditableTarget,
  shouldHandleGridSurfaceKey,
  shouldHandleGridWindowKey,
} from '../gridClipboardKeyboardController.js'
import { useWorkbookGridKeyboardHandler } from '../useWorkbookGridKeyboardHandler.js'
import { describe, expect, test, vi } from 'vitest'
import type { GridEngineLike } from '../grid-engine.js'

function createCellSnapshot(address: string, input: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    input,
    value: { tag: ValueTag.String, value: input, stringId: 0 },
    flags: 0,
    version: 0,
  }
}

function createEngine(cells: Record<string, string>): GridEngineLike {
  return {
    getCell: (_sheetName, address) => createCellSnapshot(address, cells[address] ?? ''),
    getCellStyle: () => undefined,
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  }
}

function KeyboardHandlerHarness(props: {
  beginSelectedEdit: ReturnType<typeof vi.fn>
  getGridSelection?: () => ReturnType<typeof createGridSelection>
  onClearCell?: ReturnType<typeof vi.fn>
  setGridSelection: ReturnType<typeof vi.fn>
  onSelectionChange: ReturnType<typeof vi.fn>
}) {
  const hostRef = { current: null as HTMLDivElement | null }
  const internalClipboardRef = { current: null }

  useWorkbookGridKeyboardHandler({
    applyClipboardValues: vi.fn(),
    beginSelectedEdit: props.beginSelectedEdit,
    captureInternalClipboardSelection: vi.fn(),
    editorValue: '',
    engine: {
      getCell: () => ({ value: { tag: ValueTag.String } }),
    },
    gridSelection: createGridSelection(1, 1),
    getGridSelection: props.getGridSelection,
    hostRef,
    internalClipboardRef,
    isEditingCell: false,
    onCancelEdit: vi.fn(),
    onClearCell: props.onClearCell ?? vi.fn(),
    onCommitEdit: vi.fn(),
    onEditorChange: vi.fn(),
    onSelectionChange: props.onSelectionChange,
    pendingClipboardCopySequenceRef: { current: 0 },
    pendingKeyboardPasteSequenceRef: { current: 0 },
    pendingTypeSeedRef: { current: null },
    selectedCell: { col: 1, row: 1 },
    setGridSelection: props.setGridSelection,
    sheetName: 'Sheet1',
    suppressNextNativePasteRef: { current: false },
    toggleSelectedBooleanCell: vi.fn(),
  })

  return createElement('div', {
    ref: (node: HTMLDivElement | null) => {
      hostRef.current = node
    },
  })
}

describe('gridClipboardKeyboardController', () => {
  test('routes external clipboard data through paste operations', () => {
    const onCopyRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = { current: null }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onPaste,
      sheetName: 'Sheet1',
      target: [2, 3],
      values: [['A', 'B']],
    })

    expect(onCopyRange).not.toHaveBeenCalled()
    expect(onPaste).toHaveBeenCalledWith('Sheet1', 'C4', [['A', 'B']])
  })

  test('routes matching internal clipboard data through copy-range operations', () => {
    const onCopyRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = {
      current: {
        sourceStartAddress: 'A1',
        sourceEndAddress: 'B2',
        signature: 'A\u001fB\u001eC\u001fD',
        plainText: 'A\tB\nC\tD',
        rowCount: 2,
        colCount: 2,
      },
    }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onPaste,
      sheetName: 'Sheet1',
      target: [3, 4],
      values: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    })

    expect(onCopyRange).toHaveBeenCalledWith('A1', 'B2', 'D5', 'E6')
    expect(onPaste).not.toHaveBeenCalled()
  })

  test('captures the selected grid range into an internal clipboard payload', () => {
    const internalClipboardRef = { current: null }

    const clipboard = captureGridClipboardSelection({
      engine: createEngine({
        A1: 'alpha',
        B1: 'beta',
        A2: 'gamma',
        B2: 'delta',
      }),
      gridSelection: {
        ...createGridSelection(0, 0),
        current: {
          cell: [0, 0],
          range: { x: 0, y: 0, width: 2, height: 2 },
          rangeStack: [],
        },
      },
      internalClipboardRef,
      sheetName: 'Sheet1',
    })

    expect(clipboard).toEqual({
      sourceStartAddress: 'A1',
      sourceEndAddress: 'B2',
      signature: 'alpha\u001fbeta\u001egamma\u001fdelta',
      plainText: 'alpha\tbeta\ngamma\tdelta',
      rowCount: 2,
      colCount: 2,
    })
    expect(internalClipboardRef.current).toEqual(clipboard)
  })

  test('captures optimistic editor seeds before the engine snapshot catches up', () => {
    const internalClipboardRef = { current: null }

    const clipboard = captureGridClipboardSelection({
      engine: createEngine({
        A1: 'alpha',
        B1: '',
        A2: '',
        B2: 'delta',
      }),
      getCellEditorSeed: (_sheetName, address) => {
        switch (address) {
          case 'B1':
            return 'beta'
          case 'A2':
            return 'gamma'
          default:
            return undefined
        }
      },
      gridSelection: {
        ...createGridSelection(0, 0),
        current: {
          cell: [0, 0],
          range: { x: 0, y: 0, width: 2, height: 2 },
          rangeStack: [],
        },
      },
      internalClipboardRef,
      sheetName: 'Sheet1',
    })

    expect(clipboard?.plainText).toBe('alpha\tbeta\ngamma\tdelta')
    expect(internalClipboardRef.current).toEqual(clipboard)
  })

  test('pastes the in-memory internal clipboard immediately while the system clipboard write is still pending', () => {
    const applyClipboardValues = vi.fn()
    const internalClipboardRef = {
      current: {
        sourceStartAddress: 'B2',
        sourceEndAddress: 'C3',
        signature: '3\u001f=B2*2\u001e4\u001f=B3*2',
        plainText: '3\t=B2*2\n4\t=B3*2',
        rowCount: 2,
        colCount: 2,
      },
    }

    handleGridKey({
      applyClipboardValues,
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'v',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createGridSelection(3, 1),
      internalClipboardRef,
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelectionChange: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 1 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 3, row: 1 },
      setGridSelection: vi.fn(),
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(applyClipboardValues).toHaveBeenCalledWith(
      [3, 1],
      [
        ['3', '=B2*2'],
        ['4', '=B3*2'],
      ],
    )
  })

  test('applies parsed paste payloads to the active selection and clears pending keyboard paste state', () => {
    const applyClipboardValues = vi.fn()
    const event = {
      clipboardData: {
        getData: (type: string) => (type === 'text/html' ? '<table><tr><td>A</td><td>B</td></tr></table>' : 'ignored'),
        setData: vi.fn(),
      },
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
    }

    handleGridPasteCapture({
      applyClipboardValues,
      event,
      gridSelection: createGridSelection(1, 2),
      pendingKeyboardPasteSequenceRef: { current: 3 },
      selectedCell: { col: 1, row: 2 },
      suppressNextNativePasteRef: { current: false },
    })

    expect(applyClipboardValues).toHaveBeenCalledWith([1, 2], [['A', 'B']])
    expect(event.preventDefault).toHaveBeenCalled()
    expect(event.stopPropagation).toHaveBeenCalled()
  })

  test('maps keyboard actions into selection updates', () => {
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'Enter',
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createGridSelection(2, 4),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelectionChange,
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 2, row: 4 },
      setGridSelection,
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(setGridSelection.mock.calls[0]?.[0]?.current?.cell).toEqual([2, 5])
    expect(onSelectionChange.mock.calls[0]?.[0]?.current?.cell).toEqual([2, 5])
  })

  test('toggles boolean cells with space instead of entering text edit mode', () => {
    const toggleSelectedBooleanCell = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: ' ',
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createGridSelection(1, 1),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => true,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelectionChange: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell,
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(toggleSelectedBooleanCell).toHaveBeenCalledTimes(1)
  })

  test('clears the current visible grid selection snapshot on Delete', () => {
    const onClearCell = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'Delete',
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 2]),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell,
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelectionChange: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(onClearCell).toHaveBeenCalledWith({
      sheetName: 'Sheet1',
      address: 'B2',
      kind: 'range',
      range: {
        startAddress: 'B2',
        endAddress: 'D3',
      },
    })
  })

  test('select-all updates the active address to A1', () => {
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'a',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createGridSelection(3, 7),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelectionChange,
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 3, row: 7 },
      setGridSelection,
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(setGridSelection).toHaveBeenCalledTimes(1)
    expect(onSelectionChange).toHaveBeenCalledTimes(1)
  })

  test('only claims global grid shortcuts when focus is on the document body', () => {
    const host = document.createElement('div')
    document.body.append(host)

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(true)

    const input = document.createElement('input')
    document.body.append(input)
    input.focus()
    expect(shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, input, host)).toBe(
      false,
    )

    const textarea = document.createElement('textarea')
    textarea.dataset['testid'] = 'cell-editor-input'
    document.body.append(textarea)
    textarea.focus()
    expect(isGridKeyboardEditableTarget(textarea)).toBe(true)
    expect(shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'r', metaKey: false, shiftKey: false }, textarea, host)).toBe(
      false,
    )
  })

  test('does not claim global grid shortcuts while a modal dialog is open', () => {
    const host = document.createElement('div')
    document.body.append(host)

    const modal = document.createElement('div')
    modal.setAttribute('aria-modal', 'true')
    modal.setAttribute('role', 'dialog')
    document.body.append(modal)

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'r', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Tab', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(false)
  })

  test('ignores globally prevented keydown events before routing them into the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const beginSelectedEdit = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: '?',
      shiftKey: true,
    })
    event.preventDefault()

    await act(async () => {
      window.dispatchEvent(event)
    })

    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('does not route Delete from editable event targets into the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const formulaInput = document.createElement('input')
    document.body.append(formulaInput)
    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Delete',
    })

    await act(async () => {
      formulaInput.dispatchEvent(event)
    })

    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes high-confidence grid shortcuts from workbook chrome without stealing button activation keys', () => {
    document.body.innerHTML = ''
    const scope = document.createElement('section')
    scope.dataset['workbookKeyboardScope'] = 'true'
    const gridHost = document.createElement('div')
    const toolbarButton = document.createElement('button')
    scope.append(gridHost, toolbarButton)
    document.body.append(scope)
    toolbarButton.focus()

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: 'c', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey(
        { altKey: false, ctrlKey: false, key: 'ArrowDown', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey(
        { altKey: true, ctrlKey: false, key: 'ArrowLeft', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: true, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'x', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
  })

  test('clears the latest selection when Delete is pressed from workbook chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const scope = document.createElement('section')
    scope.dataset['workbookKeyboardScope'] = 'true'
    const toolbarButton = document.createElement('button')
    const host = document.createElement('div')
    scope.append(toolbarButton, host)
    document.body.append(scope)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          getGridSelection: () => createGridSelection(4, 8),
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    toolbarButton.focus()
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Delete' }))
    })

    expect(onClearCell).toHaveBeenCalledWith({
      address: 'E9',
      kind: 'cell',
      range: { startAddress: 'E9', endAddress: 'E9' },
      sheetName: 'Sheet1',
    })
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('does not route modified delete keys into the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const events = [
      { key: 'Delete', ctrlKey: true },
      { key: 'Delete', metaKey: true },
      { key: 'Delete', altKey: true },
      { key: 'Delete', shiftKey: true },
      { key: 'Backspace', ctrlKey: true },
      { key: 'Backspace', metaKey: true },
      { key: 'Backspace', altKey: true },
      { key: 'Backspace', shiftKey: true },
    ].map(
      (eventInit) =>
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          ...eventInit,
        }),
    )

    await act(async () => {
      for (const event of events) {
        window.dispatchEvent(event)
      }
    })

    for (const event of events) {
      expect(event.defaultPrevented).toBe(false)
    }

    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('reads the latest runtime selection when routing global clear keys', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          getGridSelection: () => createGridSelection(2, 3),
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Delete',
    })

    await act(async () => {
      window.dispatchEvent(event)
    })

    expect(onClearCell).toHaveBeenCalledWith({
      address: 'C4',
      kind: 'cell',
      range: { startAddress: 'C4', endAddress: 'C4' },
      sheetName: 'Sheet1',
    })

    await act(async () => {
      root.unmount()
    })
  })

  test('filters grid-surface key handling to grid-relevant keys', () => {
    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: 'Enter',
        metaKey: false,
      }),
    ).toBe(true)

    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: 'Shift',
        metaKey: false,
      }),
    ).toBe(false)

    expect(
      shouldHandleGridSurfaceKey({
        altKey: true,
        ctrlKey: false,
        key: 'ArrowLeft',
        metaKey: false,
      }),
    ).toBe(false)
  })
})
