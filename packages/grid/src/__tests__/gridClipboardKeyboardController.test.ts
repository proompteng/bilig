// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { createColumnSliceSelection, createGridSelection, createRangeSelection, createRowSliceSelection } from '../gridSelection.js'
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  handleGridKey,
  handleGridPasteCapture,
  isGridKeyboardEditableTarget,
  shouldHandleGridSurfaceKey,
  shouldHandleGridWindowKey,
  shouldSuppressWorkbookChromeClearKey,
  shouldSuppressWorkbookChromeSelectionKeyUp,
} from '../gridClipboardKeyboardController.js'
import { createDeferredBeginEditScheduler, useWorkbookGridKeyboardHandler } from '../useWorkbookGridKeyboardHandler.js'
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

function createFormulaSnapshot(address: string, formula: string, value: number): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    formula,
    input: null,
    value: { tag: ValueTag.Number, value },
    flags: 0,
    version: 0,
  }
}

function createEngine(cells: Record<string, string | CellSnapshot>): GridEngineLike {
  return {
    getCell: (_sheetName, address) => {
      const cell = cells[address]
      return typeof cell === 'object' ? cell : createCellSnapshot(address, cell ?? '')
    },
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
  onCancelEdit?: ReturnType<typeof vi.fn>
  onClearCell?: ReturnType<typeof vi.fn>
  onCommitEdit?: ReturnType<typeof vi.fn>
  onFillRange?: ReturnType<typeof vi.fn>
  scrollActiveCellIntoView?: ReturnType<typeof vi.fn>
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
      getCell: (_sheetName, address) => createCellSnapshot(address, 'value'),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    },
    gridSelection: createGridSelection(1, 1),
    getGridSelection: props.getGridSelection,
    hostRef,
    internalClipboardRef,
    isEditingCell: false,
    onCancelEdit: props.onCancelEdit ?? vi.fn(),
    onClearCell: props.onClearCell ?? vi.fn(),
    onCommitEdit: props.onCommitEdit ?? vi.fn(),
    onEditorChange: vi.fn(),
    onFillRange: props.onFillRange ?? vi.fn(),
    onSelectionChange: props.onSelectionChange,
    scrollActiveCellIntoView: props.scrollActiveCellIntoView ?? vi.fn(),
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
  test('batches rapid typed edit seeds into one deferred begin-edit', () => {
    const beginSelectedEdit = vi.fn()
    const callbacks: FrameRequestCallback[] = []
    const scheduler = createDeferredBeginEditScheduler({
      beginSelectedEdit,
      requestAnimationFrame: (callback) => {
        callbacks.push(callback)
        return callbacks.length
      },
      cancelAnimationFrame: vi.fn(),
    })

    scheduler.schedule('a', 'caret-end')
    scheduler.schedule('ab', 'caret-end')
    scheduler.schedule('abc', 'caret-end')

    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(callbacks).toHaveLength(1)

    callbacks[0]?.(performance.now())

    expect(beginSelectedEdit).toHaveBeenCalledTimes(1)
    expect(beginSelectedEdit).toHaveBeenCalledWith('abc', 'caret-end')
  })

  test('immediate begin-edit cancels a deferred typed seed', () => {
    const beginSelectedEdit = vi.fn()
    const cancelAnimationFrame = vi.fn()
    const callbacks: FrameRequestCallback[] = []
    const scheduler = createDeferredBeginEditScheduler({
      beginSelectedEdit,
      requestAnimationFrame: (callback) => {
        callbacks.push(callback)
        return 17
      },
      cancelAnimationFrame,
    })

    scheduler.schedule('a', 'caret-end')
    scheduler.beginImmediate(undefined, 'caret-end')

    expect(cancelAnimationFrame).toHaveBeenCalledWith(17)
    expect(beginSelectedEdit).toHaveBeenCalledTimes(1)
    expect(beginSelectedEdit).toHaveBeenCalledWith(undefined, 'caret-end')

    callbacks[0]?.(performance.now())

    expect(beginSelectedEdit).toHaveBeenCalledTimes(1)
  })

  test.each([
    { expectedCommitMovement: [0, 1] as const, expectedSeed: 'ab', key: 'Enter' },
    { expectedCommitMovement: [1, 0] as const, expectedSeed: 'ab', key: 'ArrowRight' },
    { expectedBeginSeed: 'a', key: 'Backspace' },
  ] as const)('resolves a rapid typed seed when $key arrives before the deferred editor opens', async (scenario) => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const originalRequestAnimationFrame = window.requestAnimationFrame
    const originalCancelAnimationFrame = window.cancelAnimationFrame
    const callbacks: FrameRequestCallback[] = []
    window.requestAnimationFrame = vi.fn((callback: FrameRequestCallback) => {
      callbacks.push(callback)
      return callbacks.length
    })
    window.cancelAnimationFrame = vi.fn()
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const onCommitEdit = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          createElement(KeyboardHandlerHarness, {
            beginSelectedEdit,
            onClearCell,
            onCommitEdit,
            onSelectionChange,
            setGridSelection,
          }),
        )
      })

      const firstKey = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'a' })
      const secondKey = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'b' })
      const resolvingKey = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: scenario.key })

      await act(async () => {
        window.dispatchEvent(firstKey)
        window.dispatchEvent(secondKey)
        window.dispatchEvent(resolvingKey)
      })

      expect(firstKey.defaultPrevented).toBe(true)
      expect(secondKey.defaultPrevented).toBe(true)
      expect(resolvingKey.defaultPrevented).toBe(true)
      expect(callbacks).toHaveLength(1)
      expect(onClearCell).not.toHaveBeenCalled()

      if ('expectedCommitMovement' in scenario) {
        expect(beginSelectedEdit).not.toHaveBeenCalled()
        expect(onCommitEdit).toHaveBeenCalledTimes(1)
        expect(onCommitEdit).toHaveBeenCalledWith(scenario.expectedCommitMovement, scenario.expectedSeed)
      } else {
        expect(onCommitEdit).not.toHaveBeenCalled()
        expect(beginSelectedEdit).toHaveBeenCalledTimes(1)
        expect(beginSelectedEdit).toHaveBeenCalledWith(scenario.expectedBeginSeed, 'caret-end')
      }

      callbacks[0]?.(performance.now())

      expect(beginSelectedEdit).toHaveBeenCalledTimes('expectedBeginSeed' in scenario ? 1 : 0)
    } finally {
      window.requestAnimationFrame = originalRequestAnimationFrame
      window.cancelAnimationFrame = originalCancelAnimationFrame
      await act(async () => {
        root.unmount()
      })
    }
  })

  test('commits a pending typed seed before pointer selection can move it to another cell', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const originalRequestAnimationFrame = window.requestAnimationFrame
    const originalCancelAnimationFrame = window.cancelAnimationFrame
    const callbacks: FrameRequestCallback[] = []
    window.requestAnimationFrame = vi.fn((callback: FrameRequestCallback) => {
      callbacks.push(callback)
      return callbacks.length
    })
    window.cancelAnimationFrame = vi.fn()
    const beginSelectedEdit = vi.fn()
    const onCommitEdit = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          createElement(KeyboardHandlerHarness, {
            beginSelectedEdit,
            onCommitEdit,
            onSelectionChange,
            setGridSelection,
          }),
        )
      })

      const firstKey = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'a' })
      const secondKey = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'b' })
      const pointerDown = new Event('pointerdown', { bubbles: true, cancelable: true })

      await act(async () => {
        window.dispatchEvent(firstKey)
        window.dispatchEvent(secondKey)
        window.dispatchEvent(pointerDown)
      })

      expect(firstKey.defaultPrevented).toBe(true)
      expect(secondKey.defaultPrevented).toBe(true)
      expect(onCommitEdit).toHaveBeenCalledTimes(1)
      expect(onCommitEdit).toHaveBeenCalledWith(undefined, 'ab')
      expect(beginSelectedEdit).not.toHaveBeenCalled()

      callbacks[0]?.(performance.now())

      expect(beginSelectedEdit).not.toHaveBeenCalled()
    } finally {
      window.requestAnimationFrame = originalRequestAnimationFrame
      window.cancelAnimationFrame = originalCancelAnimationFrame
      await act(async () => {
        root.unmount()
      })
    }
  })

  test('routes external clipboard data through paste operations', () => {
    const onCopyRange = vi.fn()
    const onMoveRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = { current: null }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onMoveRange,
      onPaste,
      sheetName: 'Sheet1',
      target: [2, 3],
      values: [['A', 'B']],
    })

    expect(onCopyRange).not.toHaveBeenCalled()
    expect(onMoveRange).not.toHaveBeenCalled()
    expect(onPaste).toHaveBeenCalledWith('Sheet1', 'C4', [['A', 'B']])
  })

  test('routes matching internal clipboard data through copy-range operations', () => {
    const onCopyRange = vi.fn()
    const onMoveRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = {
      current: {
        operation: 'copy' as const,
        sourceStartAddress: 'A1',
        sourceEndAddress: 'B2',
        signature: 'A\u001fB\u001eC\u001fD',
        plainText: 'A\tB\nC\tD',
        valuesOnlyPlainText: 'A\tB\nC\tD',
        rowCount: 2,
        colCount: 2,
      },
    }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onMoveRange,
      onPaste,
      sheetName: 'Sheet1',
      target: [3, 4],
      values: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    })

    expect(onCopyRange).toHaveBeenCalledWith('A1', 'B2', 'D5', 'E6')
    expect(onMoveRange).not.toHaveBeenCalled()
    expect(onPaste).not.toHaveBeenCalled()
  })

  test('routes values-only paste through plain paste even when it matches an internal copied range', () => {
    const onCopyRange = vi.fn()
    const onMoveRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = {
      current: {
        operation: 'copy' as const,
        sourceStartAddress: 'B2',
        sourceEndAddress: 'C2',
        signature: '3\u001f=B2*2',
        plainText: '3\t=B2*2',
        valuesOnlyPlainText: '3\t6',
        rowCount: 1,
        colCount: 2,
      },
    }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onMoveRange,
      onPaste,
      pasteValuesOnly: true,
      sheetName: 'Sheet1',
      target: [3, 1],
      values: [['3', '6']],
    })

    expect(onCopyRange).not.toHaveBeenCalled()
    expect(onMoveRange).not.toHaveBeenCalled()
    expect(onPaste).toHaveBeenCalledWith('Sheet1', 'D2', [['3', '6']])
  })

  test('routes matching cut clipboard data through move-range operations and consumes the cut', () => {
    const onCopyRange = vi.fn()
    const onMoveRange = vi.fn()
    const onPaste = vi.fn()
    const internalClipboardRef = {
      current: {
        operation: 'cut' as const,
        sourceStartAddress: 'A1',
        sourceEndAddress: 'B2',
        signature: 'A\u001fB\u001eC\u001fD',
        plainText: 'A\tB\nC\tD',
        valuesOnlyPlainText: 'A\tB\nC\tD',
        rowCount: 2,
        colCount: 2,
      },
    }

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onMoveRange,
      onPaste,
      sheetName: 'Sheet1',
      target: [3, 4],
      values: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    })

    expect(onMoveRange).toHaveBeenCalledWith('A1', 'B2', 'D5', 'E6')
    expect(onCopyRange).not.toHaveBeenCalled()
    expect(onPaste).not.toHaveBeenCalled()
    expect(internalClipboardRef.current).toBeNull()
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
      operation: 'copy',
      sourceStartAddress: 'A1',
      sourceEndAddress: 'B2',
      signature: 'alpha\u001fbeta\u001egamma\u001fdelta',
      plainText: 'alpha\tbeta\ngamma\tdelta',
      valuesOnlyPlainText: 'alpha\tbeta\ngamma\tdelta',
      rowCount: 2,
      colCount: 2,
    })
    expect(internalClipboardRef.current).toEqual(clipboard)
  })

  test('captures formula text separately from resolved values for paste-values-only', () => {
    const internalClipboardRef = { current: null }

    const clipboard = captureGridClipboardSelection({
      engine: createEngine({
        B2: createCellSnapshot('B2', '3'),
        C2: createFormulaSnapshot('C2', 'B2*2', 6),
      }),
      gridSelection: {
        ...createGridSelection(1, 1),
        current: {
          cell: [1, 1],
          range: { x: 1, y: 1, width: 2, height: 1 },
          rangeStack: [],
        },
      },
      internalClipboardRef,
      sheetName: 'Sheet1',
    })

    expect(clipboard?.plainText).toBe('3\t=B2*2')
    expect(clipboard?.valuesOnlyPlainText).toBe('3\t6')
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
        operation: 'copy' as const,
        sourceStartAddress: 'B2',
        sourceEndAddress: 'C3',
        signature: '3\u001f=B2*2\u001e4\u001f=B3*2',
        plainText: '3\t=B2*2\n4\t=B3*2',
        valuesOnlyPlainText: '3\t6\n4\t8',
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
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
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
      { pasteValuesOnly: false },
    )
  })

  test('pastes resolved values from the internal clipboard for the paste-values-only shortcut', () => {
    const applyClipboardValues = vi.fn()
    const internalClipboardRef = {
      current: {
        operation: 'copy' as const,
        sourceStartAddress: 'B2',
        sourceEndAddress: 'C3',
        signature: '3\u001f=B2*2\u001e4\u001f=B3*2',
        plainText: '3\t=B2*2\n4\t=B3*2',
        valuesOnlyPlainText: '3\t6\n4\t8',
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
        shiftKey: true,
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
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
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
        ['3', '6'],
        ['4', '8'],
      ],
      { pasteValuesOnly: true },
    )
  })

  test('captures keyboard cut intent instead of downgrading it to copy', () => {
    const captureInternalClipboardSelection = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection,
      editorValue: '',
      event: {
        key: 'x',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createGridSelection(1, 1),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(captureInternalClipboardSelection).toHaveBeenCalledWith('cut')
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
      onFillRange: vi.fn(),
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
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
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
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
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
      onFillRange: vi.fn(),
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

  test('selects the current data region before falling back to full-sheet select-all', () => {
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
      gridSelection: createGridSelection(2, 4),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      navigation: {
        resolveCurrentRegion: () => ({ x: 1, y: 2, width: 4, height: 6 }),
        resolveDataEdge: () => null,
      },
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange,
      scrollActiveCellIntoView: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 2, row: 4 },
      setGridSelection,
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(setGridSelection).toHaveBeenCalledWith({
      columns: expect.objectContaining({ length: 0 }),
      current: {
        cell: [2, 4],
        range: { x: 1, y: 2, width: 4, height: 6 },
        rangeStack: [],
      },
      rows: expect.objectContaining({ length: 0 }),
    })
    expect(onSelectionChange).toHaveBeenCalledWith({
      columns: expect.objectContaining({ length: 0 }),
      current: {
        cell: [2, 4],
        range: { x: 1, y: 2, width: 4, height: 6 },
        rangeStack: [],
      },
      rows: expect.objectContaining({ length: 0 }),
    })
  })

  test('keeps rectangular range ownership while Enter and Tab move the active cell', () => {
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const rangeSelection = createRangeSelection(createGridSelection(1, 1), [1, 1], [2, 2])

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'Tab',
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: rangeSelection,
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange,
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection,
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(setGridSelection).toHaveBeenCalledWith({
      columns: expect.objectContaining({ length: 0 }),
      current: {
        cell: [2, 1],
        range: { x: 1, y: 1, width: 2, height: 2 },
        rangeStack: [],
      },
      rows: expect.objectContaining({ length: 0 }),
    })
    expect(onSelectionChange).toHaveBeenCalledWith({
      columns: expect.objectContaining({ length: 0 }),
      current: {
        cell: [2, 1],
        range: { x: 1, y: 1, width: 2, height: 2 },
        rangeStack: [],
      },
      rows: expect.objectContaining({ length: 0 }),
    })
  })

  test('routes fill down and fill right keyboard shortcuts through range fill operations', () => {
    const onFillRange = vi.fn()
    const fillDownEvent = {
      key: 'd',
      ctrlKey: true,
      metaKey: false,
      altKey: false,
      preventDefault: vi.fn(),
    }
    const fillRightEvent = {
      key: 'r',
      ctrlKey: false,
      metaKey: true,
      altKey: false,
      preventDefault: vi.fn(),
    }

    for (const event of [fillDownEvent, fillRightEvent]) {
      handleGridKey({
        applyClipboardValues: vi.fn(),
        beginSelectedEdit: vi.fn(),
        captureInternalClipboardSelection: vi.fn(),
        editorValue: '',
        event,
        gridSelection: createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 4]),
        internalClipboardRef: { current: null },
        isSelectedCellBoolean: () => false,
        isEditingCell: false,
        onCancelEdit: vi.fn(),
        onClearCell: vi.fn(),
        onCommitEdit: vi.fn(),
        onEditorChange: vi.fn(),
        onFillRange,
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
    }

    expect(fillDownEvent.preventDefault).toHaveBeenCalled()
    expect(fillRightEvent.preventDefault).toHaveBeenCalled()
    expect(onFillRange).toHaveBeenNthCalledWith(1, 'B2', 'D2', 'B3', 'D5')
    expect(onFillRange).toHaveBeenNthCalledWith(2, 'B2', 'B5', 'C2', 'D5')
  })

  test('routes structural delete shortcut through selected row and column mutations', () => {
    const onDeleteRows = vi.fn()
    const onDeleteColumns = vi.fn()
    const rowDeleteEvent = {
      key: '-',
      ctrlKey: true,
      metaKey: false,
      altKey: true,
      preventDefault: vi.fn(),
      cancel: vi.fn(),
    }
    const columnDeleteEvent = {
      key: '-',
      ctrlKey: false,
      metaKey: true,
      altKey: true,
      preventDefault: vi.fn(),
      cancel: vi.fn(),
    }

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: rowDeleteEvent,
      gridSelection: createRowSliceSelection(0, 2, 4),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onDeleteRows,
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 0, row: 2 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: columnDeleteEvent,
      gridSelection: createColumnSliceSelection(1, 3, 0),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onDeleteColumns,
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 0 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(rowDeleteEvent.preventDefault).toHaveBeenCalled()
    expect(rowDeleteEvent.cancel).toHaveBeenCalled()
    expect(columnDeleteEvent.preventDefault).toHaveBeenCalled()
    expect(columnDeleteEvent.cancel).toHaveBeenCalled()
    expect(onDeleteRows).toHaveBeenCalledWith(2, 3)
    expect(onDeleteColumns).toHaveBeenCalledWith(1, 3)
  })

  test('routes primary-modified Backspace to active-cell scrolling without clearing content', () => {
    const onClearCell = vi.fn()
    const scrollActiveCellIntoView = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'Backspace',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createGridSelection(3, 7),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell,
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView,
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 3, row: 7 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(scrollActiveCellIntoView).toHaveBeenCalledTimes(1)
    expect(onClearCell).not.toHaveBeenCalled()
  })

  test('suppresses no-op fill shortcuts so the browser does not steal grid focus', () => {
    const onFillRange = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'd',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 1]),
      internalClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange,
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(onFillRange).not.toHaveBeenCalled()
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

  test('routes high-confidence grid shortcuts from workbook chrome without stealing button activation keys or delete ownership', () => {
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
    ).toBe(false)
    expect(
      shouldSuppressWorkbookChromeClearKey(
        { altKey: false, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldSuppressWorkbookChromeClearKey(
        { altKey: false, ctrlKey: false, key: 'Backspace', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: 'c', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: true, ctrlKey: true, key: '-', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: true }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: true }, toolbarButton, gridHost),
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
      shouldSuppressWorkbookChromeClearKey(
        { altKey: true, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
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
    expect(
      shouldSuppressWorkbookChromeSelectionKeyUp(
        { altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: true },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldSuppressWorkbookChromeSelectionKeyUp(
        { altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(false)
  })

  test('suppresses browser clear-key defaults from workbook chrome without clearing the grid', async () => {
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
    const deleteEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Delete' })
    const backspaceEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Backspace' })
    await act(async () => {
      toolbarButton.dispatchEvent(deleteEvent)
      toolbarButton.dispatchEvent(backspaceEvent)
    })

    expect(deleteEvent.defaultPrevented).toBe(true)
    expect(backspaceEvent.defaultPrevented).toBe(true)
    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes row, column, and full-sheet selection shortcuts from workbook chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
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
          getGridSelection: () => createGridSelection(3, 7),
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    toolbarButton.focus()
    const rowEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: ' ', shiftKey: true })
    const columnEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: true, key: ' ' })
    const sheetEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: true, key: ' ', shiftKey: true })
    await act(async () => {
      toolbarButton.dispatchEvent(rowEvent)
      toolbarButton.dispatchEvent(columnEvent)
      toolbarButton.dispatchEvent(sheetEvent)
    })

    expect(rowEvent.defaultPrevented).toBe(true)
    expect(columnEvent.defaultPrevented).toBe(true)
    expect(sheetEvent.defaultPrevented).toBe(true)
    expect(setGridSelection).toHaveBeenCalledTimes(3)
    expect(setGridSelection.mock.calls[0]?.[0]?.rows.first()).toBe(7)
    expect(setGridSelection.mock.calls[0]?.[0]?.columns.first()).toBeUndefined()
    expect(setGridSelection.mock.calls[1]?.[0]?.columns.first()).toBe(3)
    expect(setGridSelection.mock.calls[1]?.[0]?.rows.first()).toBeUndefined()
    expect(setGridSelection.mock.calls[2]?.[0]?.current?.cell).toEqual([0, 0])
    expect(setGridSelection.mock.calls[2]?.[0]?.rows.first()).toBe(0)
    expect(setGridSelection.mock.calls[2]?.[0]?.columns.first()).toBe(0)
    expect(onSelectionChange).toHaveBeenCalledTimes(3)
    expect(beginSelectedEdit).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes only primary Backspace among modified delete keys', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const scrollActiveCellIntoView = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onClearCell,
          onSelectionChange,
          scrollActiveCellIntoView,
          setGridSelection,
        }),
      )
    })

    const events = [
      { key: 'Delete', ctrlKey: true },
      { key: 'Delete', metaKey: true },
      { key: 'Delete', altKey: true },
      { key: 'Delete', shiftKey: true },
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
    expect(scrollActiveCellIntoView).not.toHaveBeenCalled()

    const primaryBackspace = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      ctrlKey: true,
      key: 'Backspace',
    })
    await act(async () => {
      window.dispatchEvent(primaryBackspace)
    })
    expect(primaryBackspace.defaultPrevented).toBe(true)
    expect(scrollActiveCellIntoView).toHaveBeenCalledTimes(1)
    expect(onClearCell).not.toHaveBeenCalled()

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
        key: '-',
        metaKey: true,
      }),
    ).toBe(true)

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
