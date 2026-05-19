// @vitest-environment jsdom
import { act, type MutableRefObject } from 'react'
import { createRoot } from 'react-dom/client'
import { WorkbookGridFocusReturnProvider } from '@bilig/grid'
import { ValueTag, type CellRangeRef, type CellStyleRecord } from '@bilig/protocol'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { getWorkbookShortcutLabel } from '../shortcut-registry.js'
import { WorkbookToolbar } from '../workbook-toolbar.js'
import { deriveWorkbookStatusPresentation, useWorkbookToolbar } from '../use-workbook-toolbar.js'

function ToolbarHookHarness(props: {
  readonly invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  readonly onRedo?: (() => void) | undefined
  readonly onUndo?: (() => void) | undefined
  readonly selectionRangeRef: MutableRefObject<CellRangeRef>
  readonly selectedStyle?: CellStyleRecord | undefined
  readonly canHideCurrentRow?: boolean | undefined
  readonly canUnmergeSelection?: boolean | undefined
  readonly writesAllowed?: boolean | undefined
}) {
  const { ribbon } = useWorkbookToolbar({
    canHideCurrentColumn: false,
    canHideCurrentRow: props.canHideCurrentRow ?? false,
    canRedo: false,
    canUndo: false,
    canUnhideCurrentColumn: false,
    canUnhideCurrentRow: false,
    canUnmergeSelection: props.canUnmergeSelection ?? false,
    connectionStateName: 'connected',
    currentFillColor: '#ffffff',
    currentNumberFormatKind: 'general',
    currentTextColor: '#111827',
    horizontalAlignment: null,
    invokeMutation: props.invokeMutation,
    localPersistenceMode: 'ephemeral',
    onApplyBorderPreset: () => {},
    onClearStyle: () => {},
    onFillColorReset: () => {},
    onFillColorSelect: () => {},
    onFontSizeChange: () => {},
    onHideCurrentColumn: () => {},
    onHideCurrentRow: () => {},
    onHorizontalAlignmentChange: () => {},
    onNumberFormatChange: () => {},
    onRedo: props.onRedo ?? (() => {}),
    onTextColorReset: () => {},
    onTextColorSelect: () => {},
    onToggleBold: () => {},
    onToggleItalic: () => {},
    onToggleUnderline: () => {},
    onToggleWrap: () => {},
    onUndo: props.onUndo ?? (() => {}),
    onUnhideCurrentColumn: () => {},
    onUnhideCurrentRow: () => {},
    remoteSyncAvailable: true,
    runtimeReady: true,
    selectedCell: {
      address: 'A1',
      sheetName: 'Sheet1',
      flags: 0,
      value: { tag: ValueTag.Empty },
      version: 0,
    },
    selectedStyle: props.selectedStyle,
    selection: { sheetName: 'Sheet1' },
    selectionRangeRef: props.selectionRangeRef,
    trailingContent: null,
    writesAllowed: props.writesAllowed ?? true,
    zeroConfigured: true,
    zeroHealthReady: true,
  })

  return <>{ribbon}</>
}

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

function setScrollGeometry(
  element: Element,
  geometry: {
    readonly clientWidth: number
    readonly scrollLeft?: number | undefined
    readonly scrollWidth: number
  },
) {
  Object.defineProperties(element, {
    clientWidth: {
      configurable: true,
      value: geometry.clientWidth,
    },
    scrollLeft: {
      configurable: true,
      value: geometry.scrollLeft ?? 0,
      writable: true,
    },
    scrollWidth: {
      configurable: true,
      value: geometry.scrollWidth,
    },
  })
}

function dispatchWorkbookShortcut(init: KeyboardEventInit & { readonly key: string }) {
  const event = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    ...init,
  })
  window.dispatchEvent(event)
  return event
}

async function flushToolbarMutationQueue(cycles = 3): Promise<void> {
  await Promise.resolve()
  if (cycles > 1) {
    await flushToolbarMutationQueue(cycles - 1)
  }
}

describe('WorkbookToolbar', () => {
  it('shows shortcut formatting state optimistically while the mutation is still pending', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let resolveMutation: (() => void) | undefined
    const invokeMutation = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          resolveMutation = resolve
        }),
    )
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
    })

    const underline = host.querySelector("[aria-label='Underline']")
    expect(invokeMutation).toHaveBeenCalledWith(
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { underline: true },
      },
    )
    expect(underline?.getAttribute('aria-pressed')).toBe('true')
    expect(underline?.className).toContain('bg-[var(--wb-accent-soft)]')

    await act(async () => {
      resolveMutation?.()
    })
    await act(async () => {
      root.unmount()
    })
  })

  it('routes supported workbook toolbar shortcuts to the active selection', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const onRedo = vi.fn()
    const onUndo = vi.fn()
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'D5',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const selectionRange = selectionRangeRef.current

    await act(async () => {
      root.render(
        <ToolbarHookHarness invokeMutation={invokeMutation} onRedo={onRedo} onUndo={onUndo} selectionRangeRef={selectionRangeRef} />,
      )
    })

    await act(async () => {
      dispatchWorkbookShortcut({ key: 'z', metaKey: true })
      dispatchWorkbookShortcut({ key: 'z', metaKey: true, shiftKey: true })
      dispatchWorkbookShortcut({ key: 'y', ctrlKey: true })
    })

    expect(onUndo).toHaveBeenCalledTimes(1)
    expect(onRedo).toHaveBeenCalledTimes(2)

    const shortcutCases: readonly {
      readonly event: KeyboardEventInit & { readonly key: string }
      readonly calls: readonly unknown[][]
    }[] = [
      {
        event: { key: 'b', metaKey: true },
        calls: [['setRangeStyle', selectionRange, { font: { bold: true } }]],
      },
      {
        event: { key: 'i', metaKey: true },
        calls: [['setRangeStyle', selectionRange, { font: { italic: true } }]],
      },
      {
        event: { key: 'u', metaKey: true },
        calls: [['setRangeStyle', selectionRange, { font: { underline: true } }]],
      },
      {
        event: { code: 'Digit1', key: '!', metaKey: true, shiftKey: true },
        calls: [['setRangeNumberFormat', selectionRange, { kind: 'number', decimals: 2, useGrouping: true }]],
      },
      {
        event: { code: 'Digit4', key: '$', metaKey: true, shiftKey: true },
        calls: [
          [
            'setRangeNumberFormat',
            selectionRange,
            {
              kind: 'currency',
              currency: 'USD',
              decimals: 2,
              useGrouping: true,
              negativeStyle: 'minus',
              zeroStyle: 'zero',
            },
          ],
        ],
      },
      {
        event: { code: 'Digit5', key: '%', metaKey: true, shiftKey: true },
        calls: [['setRangeNumberFormat', selectionRange, { kind: 'percent', decimals: 2 }]],
      },
      {
        event: { key: 'l', metaKey: true, shiftKey: true },
        calls: [['setRangeStyle', selectionRange, { alignment: { horizontal: 'left' } }]],
      },
      {
        event: { key: 'e', metaKey: true, shiftKey: true },
        calls: [['setRangeStyle', selectionRange, { alignment: { horizontal: 'center' } }]],
      },
      {
        event: { key: 'r', metaKey: true, shiftKey: true },
        calls: [['setRangeStyle', selectionRange, { alignment: { horizontal: 'right' } }]],
      },
    ]

    await act(async () => {
      for (const shortcutCase of shortcutCases) {
        const event = dispatchWorkbookShortcut(shortcutCase.event)
        expect(event.defaultPrevented).toBe(true)
      }
      await flushToolbarMutationQueue(30)
    })
    expect(invokeMutation.mock.calls).toEqual(shortcutCases.flatMap((shortcutCase) => shortcutCase.calls))

    invokeMutation.mockClear()
    await act(async () => {
      const event = dispatchWorkbookShortcut({ code: 'Backslash', key: '\\', metaKey: true })
      expect(event.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })
    expect(invokeMutation.mock.calls).toEqual([
      ['clearRangeStyle', selectionRange, undefined],
      ['clearRangeNumberFormat', selectionRange],
    ])

    invokeMutation.mockClear()
    await act(async () => {
      const event = dispatchWorkbookShortcut({ code: 'Digit7', key: '&', metaKey: true, shiftKey: true })
      expect(event.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })
    const borderTrigger = host.querySelector("[aria-label='Borders']")
    expect(borderTrigger?.getAttribute('aria-pressed')).toBe('true')
    expect(borderTrigger?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation.mock.calls).toEqual([
      ['clearRangeStyle', selectionRange, ['borderTop', 'borderRight', 'borderBottom', 'borderLeft']],
      [
        'setRangeStyle',
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'D2' },
        { borders: { top: { color: '#111827', style: 'solid', weight: 'thin' } } },
      ],
      [
        'setRangeStyle',
        { sheetName: 'Sheet1', startAddress: 'B5', endAddress: 'D5' },
        { borders: { bottom: { color: '#111827', style: 'solid', weight: 'thin' } } },
      ],
      [
        'setRangeStyle',
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B5' },
        { borders: { left: { color: '#111827', style: 'solid', weight: 'thin' } } },
      ],
      [
        'setRangeStyle',
        { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D5' },
        { borders: { right: { color: '#111827', style: 'solid', weight: 'thin' } } },
      ],
    ])

    await act(async () => {
      root.unmount()
    })
  })

  it('serializes border shortcut clears before applying replacement borders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    let resolveClear: (() => void) | undefined
    const invokeMutation = vi.fn((method: string) => {
      if (method === 'clearRangeStyle') {
        return new Promise<void>((resolve) => {
          resolveClear = resolve
        })
      }
      return Promise.resolve()
    })
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      const event = dispatchWorkbookShortcut({ code: 'Digit7', key: '&', metaKey: true, shiftKey: true })
      expect(event.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })

    expect(invokeMutation.mock.calls).toEqual([
      ['clearRangeStyle', selectionRangeRef.current, ['borderTop', 'borderRight', 'borderBottom', 'borderLeft']],
    ])
    expect(host.querySelector("[aria-label='Borders']")?.getAttribute('aria-pressed')).toBe('true')

    await act(async () => {
      resolveClear?.()
      await flushToolbarMutationQueue()
    })

    expect(invokeMutation.mock.calls.slice(1)).toEqual([
      ['setRangeStyle', selectionRangeRef.current, { borders: { top: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { bottom: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { left: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { right: { color: '#111827', style: 'solid', weight: 'thin' } } }],
    ])

    await act(async () => {
      root.unmount()
    })
  })

  it('serializes clear-style shortcuts after pending border shortcut writes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const pendingBorderWrites: Array<() => void> = []
    const invokeMutation = vi.fn((method: string) => {
      if (method === 'setRangeStyle') {
        return new Promise<void>((resolve) => {
          pendingBorderWrites.push(resolve)
        })
      }
      return Promise.resolve()
    })
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      const event = dispatchWorkbookShortcut({ code: 'Digit7', key: '&', metaKey: true, shiftKey: true })
      expect(event.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })

    expect(invokeMutation.mock.calls).toEqual([
      ['clearRangeStyle', selectionRangeRef.current, ['borderTop', 'borderRight', 'borderBottom', 'borderLeft']],
      ['setRangeStyle', selectionRangeRef.current, { borders: { top: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { bottom: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { left: { color: '#111827', style: 'solid', weight: 'thin' } } }],
      ['setRangeStyle', selectionRangeRef.current, { borders: { right: { color: '#111827', style: 'solid', weight: 'thin' } } }],
    ])
    expect(pendingBorderWrites).toHaveLength(4)

    await act(async () => {
      const event = dispatchWorkbookShortcut({ code: 'Backslash', key: '\\', metaKey: true })
      expect(event.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })

    expect(invokeMutation.mock.calls).toHaveLength(5)
    expect(host.querySelector("[aria-label='Borders']")?.getAttribute('aria-pressed')).toBe('false')

    await act(async () => {
      pendingBorderWrites.forEach((resolve) => resolve())
      await flushToolbarMutationQueue()
    })

    expect(invokeMutation.mock.calls.slice(5)).toEqual([
      ['clearRangeStyle', selectionRangeRef.current, undefined],
      ['clearRangeNumberFormat', selectionRangeRef.current],
    ])

    await act(async () => {
      root.unmount()
    })
  })

  it('routes native browser history input events to workbook undo and redo', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const onRedo = vi.fn()
    const onUndo = vi.fn()
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ToolbarHookHarness invokeMutation={invokeMutation} onRedo={onRedo} onUndo={onUndo} selectionRangeRef={selectionRangeRef} />,
      )
    })

    await act(async () => {
      const undoEvent = new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'historyUndo' })
      const redoEvent = new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'historyRedo' })
      window.dispatchEvent(undoEvent)
      window.dispatchEvent(redoEvent)
      expect(undoEvent.defaultPrevented).toBe(true)
      expect(redoEvent.defaultPrevented).toBe(true)
      await flushToolbarMutationQueue()
    })

    expect(onUndo).toHaveBeenCalledTimes(1)
    expect(onRedo).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps workbook history shortcuts inside active text entry targets', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const onRedo = vi.fn()
    const onUndo = vi.fn()
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const textInput = document.createElement('input')
    textInput.value = 'native-history'
    document.body.appendChild(textInput)

    await act(async () => {
      root.render(
        <ToolbarHookHarness invokeMutation={invokeMutation} onRedo={onRedo} onUndo={onUndo} selectionRangeRef={selectionRangeRef} />,
      )
    })

    await act(async () => {
      const undoKeyEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'z', metaKey: true })
      const redoInputEvent = new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'historyRedo' })
      textInput.dispatchEvent(undoKeyEvent)
      textInput.dispatchEvent(redoInputEvent)
      expect(undoKeyEvent.defaultPrevented).toBe(false)
      expect(redoInputEvent.defaultPrevented).toBe(false)
    })

    expect(onUndo).not.toHaveBeenCalled()
    expect(onRedo).not.toHaveBeenCalled()

    textInput.remove()
    await act(async () => {
      root.unmount()
    })
  })

  it('derives clear visible save states for saved, saving, local-only, offline, and sync issues', () => {
    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 0, failedCount: 0 },
        remoteSyncAvailable: true,
        zeroConfigured: true,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toEqual({
      modeLabel: 'Live',
      syncLabel: 'Saved',
      tone: 'positive',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 2, failedCount: 0 },
        remoteSyncAvailable: true,
        zeroConfigured: true,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Sync pending',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        hasLocalMutationInFlight: true,
        pendingMutationSummary: { activeCount: 0, failedCount: 0 },
        remoteSyncAvailable: true,
        zeroConfigured: true,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Sync pending',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connecting',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 2, failedCount: 0 },
        remoteSyncAvailable: false,
        zeroConfigured: true,
        zeroHealthReady: false,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Sync pending',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connecting',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 0, failedCount: 0 },
        remoteSyncAvailable: false,
        zeroConfigured: true,
        zeroHealthReady: false,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Local saved',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 0, failedCount: 0 },
        remoteSyncAvailable: true,
        zeroConfigured: false,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Local only',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'closed',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 2, failedCount: 0 },
        remoteSyncAvailable: false,
        zeroConfigured: false,
        zeroHealthReady: false,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Local saved',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'disconnected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 0, failedCount: 0 },
        remoteSyncAvailable: false,
        zeroConfigured: true,
        zeroHealthReady: false,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Offline',
      tone: 'warning',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'ephemeral',
        pendingMutationSummary: { activeCount: 0, failedCount: 1 },
        failedPendingMutation: { id: 'pending-1' },
        remoteSyncAvailable: true,
        zeroConfigured: true,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Sync issue',
      tone: 'danger',
    })
  })

  it('does not render a domain-specific templates menu', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    expect(document.querySelector("[aria-label='Templates']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('opens the structure menu and invokes current row actions', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const onHideCurrentRow = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={onHideCurrentRow}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      )
    })

    const trigger = document.querySelector("[aria-label='Structure']")
    expect(trigger).not.toBeNull()

    await act(async () => {
      trigger?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const hideRowButton = document.querySelector("[aria-label='Hide row']")
    expect(hideRowButton).not.toBeNull()
    expect(hideRowButton?.getAttribute('disabled')).toBeNull()
    expect(hideRowButton?.getAttribute('class')).toContain('text-[var(--color-mauve-900)]')

    const unhideRowButton = document.querySelector("[aria-label='Unhide row']")
    expect(unhideRowButton).not.toBeNull()
    expect(unhideRowButton?.getAttribute('class')).toContain('text-[var(--wb-text-muted)]')
    expect(unhideRowButton?.getAttribute('class')).toContain('opacity-45')

    await act(async () => {
      hideRowButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(onHideCurrentRow).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('shows shared shortcut labels on alignment buttons', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment="left"
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      )
    })

    const alignLeftButton = document.querySelector("[aria-label='Align left']")
    const alignCenterButton = document.querySelector("[aria-label='Align center']")
    const alignRightButton = document.querySelector("[aria-label='Align right']")

    expect(alignLeftButton?.getAttribute('title')).toBe(`Align left (${getWorkbookShortcutLabel('align-left')})`)
    expect(alignCenterButton?.getAttribute('title')).toBe(`Align center (${getWorkbookShortcutLabel('align-center')})`)
    expect(alignRightButton?.getAttribute('title')).toBe(`Align right (${getWorkbookShortcutLabel('align-right')})`)

    await act(async () => {
      root.unmount()
    })
  })

  it('returns grid focus after direct toolbar and transient palette commands', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const requestGridFocus = vi.fn()
    const onFillColorSelect = vi.fn()
    const onToggleBold = vi.fn()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const flushFocusReturn = async () => {
      await act(async () => {
        await new Promise((resolve) => window.setTimeout(resolve, 0))
      })
    }

    await act(async () => {
      root.render(
        <WorkbookGridFocusReturnProvider requestGridFocus={requestGridFocus}>
          <WorkbookToolbar
            canHideCurrentColumn={false}
            canHideCurrentRow={false}
            canMergeSelection={false}
            canUnmergeSelection={false}
            canRedo={false}
            canUndo={false}
            canUnhideCurrentColumn={false}
            canUnhideCurrentRow={false}
            currentFillColor="#ffffff"
            currentNumberFormatKind="general"
            currentTextColor="#111827"
            horizontalAlignment="left"
            isBoldActive={false}
            isItalicActive={false}
            isUnderlineActive={false}
            isWrapActive={false}
            onApplyBorderPreset={() => {}}
            onClearStyle={() => {}}
            onFillColorReset={() => {}}
            onFillColorSelect={onFillColorSelect}
            onFontSizeChange={() => {}}
            onHideCurrentColumn={() => {}}
            onHideCurrentRow={() => {}}
            onMergeSelectedCells={() => {}}
            onHorizontalAlignmentChange={() => {}}
            onNumberFormatChange={() => {}}
            onRedo={() => {}}
            onTextColorReset={() => {}}
            onTextColorSelect={() => {}}
            onToggleBold={onToggleBold}
            onToggleItalic={() => {}}
            onToggleUnderline={() => {}}
            onToggleWrap={() => {}}
            onUndo={() => {}}
            onUnmergeSelectedCells={() => {}}
            onUnhideCurrentColumn={() => {}}
            onUnhideCurrentRow={() => {}}
            recentFillColors={[]}
            recentTextColors={[]}
            selectedFontSize="11"
            writesAllowed
          />
        </WorkbookGridFocusReturnProvider>,
      )
    })

    await act(async () => {
      host.querySelector("[aria-label='Bold']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onToggleBold).toHaveBeenCalledTimes(1)

    await flushFocusReturn()
    expect(requestGridFocus).toHaveBeenCalledTimes(1)

    await act(async () => {
      host.querySelector("[aria-label='Fill color']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const greenSwatch = document.querySelector("[aria-label='Fill color green']")
    expect(greenSwatch).not.toBeNull()

    await act(async () => {
      greenSwatch?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(onFillColorSelect).toHaveBeenCalledWith('#00ff00', 'preset')

    await flushFocusReturn()
    expect(requestGridFocus).toHaveBeenCalledTimes(2)

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps toolbar controls on one shared height system', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo
          canUndo
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment="left"
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          writesAllowed
        />,
      )
    })

    const historyGroup = host.querySelector("[aria-label='History']")
    const undoButton = host.querySelector("[aria-label='Undo']")
    const numberFormatTrigger = host.querySelector("[aria-label='Number format']")
    const fontSizeTrigger = host.querySelector("[aria-label='Font size']")
    const structureTrigger = host.querySelector("[aria-label='Structure']")

    expect(historyGroup?.className).toContain('h-8')
    expect(undoButton?.className).toContain('h-8')
    expect(numberFormatTrigger?.className).toContain('h-8')
    expect(numberFormatTrigger?.className).toContain('max-[420px]:w-28')
    expect(fontSizeTrigger?.className).toContain('h-8')
    expect(fontSizeTrigger?.className).toContain('max-[420px]:w-14')
    expect(structureTrigger?.className).toContain('h-8')
    expect(structureTrigger?.textContent).not.toContain('Structure')
    expect(structureTrigger?.querySelector('svg')).not.toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it("renders trailing controls in the toolbar's right slot", async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div data-testid="toolbar-tail-probe">Tail</div>}
          writesAllowed
        />,
      )
    })

    const trailingSlot = host.querySelector("[data-testid='toolbar-trailing-content']")
    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    expect(trailingSlot).not.toBeNull()
    expect(trailingSlot?.className).toContain('ml-auto')
    expect(trailingSlot?.className).toContain('flex-none')
    expect(formattingScroll?.className).toContain('flex-1')
    expect(formattingScroll?.className).toContain('min-w-0')
    expect(trailingSlot?.textContent).toContain('Tail')

    await act(async () => {
      root.unmount()
    })
  })

  it('targets the live selection range for formatting actions without waiting for a rerender', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    selectionRangeRef.current = {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    }

    await act(async () => {
      host.querySelector("[aria-label='Bold']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(invokeMutation).toHaveBeenCalledWith(
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'D5',
      },
      {
        font: { bold: true },
      },
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('dispatches merge and unmerge mutations from the structure menu', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'D5',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      document.querySelector("[aria-label='Structure']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    await act(async () => {
      document.querySelector("[aria-label='Merge cells']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(document.querySelector("[aria-label='Merge cells']")).toBeNull()
    await act(async () => {
      root.render(<ToolbarHookHarness canUnmergeSelection invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })
    await act(async () => {
      document.querySelector("[aria-label='Structure']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await act(async () => {
      document.querySelector("[aria-label='Unmerge cells']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(document.querySelector("[aria-label='Unmerge cells']")).toBeNull()

    expect(invokeMutation).toHaveBeenCalledWith('mergeCells', {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    })
    expect(invokeMutation).toHaveBeenCalledWith('unmergeCells', {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps unmerge disabled until the visible selection intersects a merged range', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness canHideCurrentRow invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      document.querySelector("[aria-label='Structure']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const unmergeButton = document.querySelector<HTMLButtonElement>("[aria-label='Unmerge cells']")
    expect(unmergeButton).not.toBeNull()
    expect(unmergeButton?.disabled).toBe(true)

    await act(async () => {
      unmergeButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(invokeMutation).not.toHaveBeenCalledWith('unmergeCells', expect.anything())

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps global formatting shortcut capture mounted across active style rerenders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const addSpy = vi.spyOn(window, 'addEventListener')
    const removeSpy = vi.spyOn(window, 'removeEventListener')
    const keydownAdds = () => addSpy.mock.calls.filter(([type]) => type === 'keydown').length
    const keydownRemoves = () => removeSpy.mock.calls.filter(([type]) => type === 'keydown').length

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(0)

    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{ id: 'style-bold', font: { bold: true } }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })
    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{ id: 'style-bold-italic', font: { bold: true, italic: true } }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(0)

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
    })

    expect(invokeMutation).toHaveBeenLastCalledWith(
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { underline: true },
      },
    )

    await act(async () => {
      root.unmount()
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(1)
  })

  it('optimistically marks formatting shortcut buttons while mutations are pending', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const pendingMutationResolvers: Array<() => void> = []
    const pendingMutations: Array<Promise<void>> = []
    const invokeMutation = vi.fn(() => {
      const pendingMutation = new Promise<void>((resolve) => {
        pendingMutationResolvers.push(resolve)
      })
      pendingMutations.push(pendingMutation)
      return pendingMutation
    })
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'b', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      1,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { bold: true },
      },
    )

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'i', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Underline']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)

    await act(async () => {
      pendingMutationResolvers[0]?.()
      await flushToolbarMutationQueue()
    })
    expect(invokeMutation).toHaveBeenCalledTimes(2)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      2,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { italic: true },
      },
    )

    await act(async () => {
      pendingMutationResolvers[1]?.()
      await flushToolbarMutationQueue()
    })
    expect(invokeMutation).toHaveBeenCalledTimes(3)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      3,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { underline: true },
      },
    )

    await act(async () => {
      pendingMutationResolvers[2]?.()
      await flushToolbarMutationQueue()
      await Promise.all(pendingMutations)
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('marks the border menu active when the selected cell has borders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{
            id: 'style-bordered',
            borders: {
              top: {
                color: '#111827',
                style: 'solid',
                weight: 'thin',
              },
            },
          }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })

    const borderTrigger = host.querySelector("[aria-label='Borders']")
    expect(borderTrigger?.getAttribute('aria-pressed')).toBe('true')
    expect(borderTrigger?.className).toContain('bg-[var(--wb-accent-soft)]')

    await act(async () => {
      root.unmount()
    })
  })

  it('applies bottom border presets to the live selection range', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    selectionRangeRef.current = {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    }

    await act(async () => {
      host.querySelector("[aria-label='Borders']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    await act(async () => {
      document.querySelector("[aria-label='Bottom border']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(invokeMutation.mock.calls).toEqual([
      [
        'clearRangeStyle',
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'D5',
        },
        ['borderTop', 'borderRight', 'borderBottom', 'borderLeft'],
      ],
      [
        'setRangeStyle',
        {
          sheetName: 'Sheet1',
          startAddress: 'B5',
          endAddress: 'D5',
        },
        {
          borders: {
            bottom: {
              style: 'solid',
              weight: 'thin',
              color: '#111827',
            },
          },
        },
      ],
    ])

    await act(async () => {
      root.unmount()
    })
  })

  it('hides native toolbar overflow scrollbars while preserving horizontal scrolling', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      )
    })

    const toolbar = host.querySelector("[aria-label='Formatting toolbar']")
    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    expect(toolbar?.className).toContain('overflow-hidden')
    expect(formattingScroll?.className).toContain('overflow-x-auto')
    expect(formattingScroll?.className).toContain('overflow-y-hidden')
    expect(formattingScroll?.className).toContain('wb-scrollbar-none')

    await act(async () => {
      root.unmount()
    })
  })

  it('signals hidden formatting actions on narrow toolbar scroll regions', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      )
    })

    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    if (!formattingScroll) {
      throw new Error('Expected formatting toolbar scroll region to render')
    }

    setScrollGeometry(formattingScroll, { clientWidth: 220, scrollWidth: 870 })
    const scrollBy = vi.fn()
    Object.defineProperty(formattingScroll, 'scrollBy', {
      configurable: true,
      value: scrollBy,
    })

    await act(async () => {
      window.dispatchEvent(new Event('resize'))
    })

    const cue = host.querySelector("[data-testid='toolbar-overflow-cue']")
    expect(host.querySelector("[data-testid='toolbar-overflow-back-cue']")).toBeNull()
    expect(cue).not.toBeNull()
    expect(cue?.getAttribute('aria-label')).toBe('Show more toolbar actions')
    expect(formattingScroll.className).not.toContain('pr-7')
    expect(cue?.className).not.toContain('absolute')
    expect(cue?.className).toContain('flex-none')
    expect(cue?.className).toContain('text-[var(--wb-accent)]')

    await act(async () => {
      cue?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(scrollBy).toHaveBeenCalledWith({
      behavior: 'smooth',
      left: 220,
    })

    setScrollGeometry(formattingScroll, {
      clientWidth: 220,
      scrollLeft: 650,
      scrollWidth: 870,
    })

    await act(async () => {
      formattingScroll?.dispatchEvent(new Event('scroll'))
    })

    expect(host.querySelector("[data-testid='toolbar-overflow-cue']")).toBeNull()
    const backCue = host.querySelector("[data-testid='toolbar-overflow-back-cue']")
    expect(backCue).not.toBeNull()
    expect(backCue?.getAttribute('aria-label')).toBe('Show previous toolbar actions')
    expect(backCue?.className).toContain('flex-none')
    expect(backCue?.className).toContain('border-r')

    await act(async () => {
      backCue?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(scrollBy).toHaveBeenLastCalledWith({
      behavior: 'smooth',
      left: -275,
    })

    await act(async () => {
      root.unmount()
    })
  })
})
