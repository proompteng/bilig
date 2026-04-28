// @vitest-environment jsdom
import { act, type MutableRefObject } from 'react'
import { createRoot } from 'react-dom/client'
import { ValueTag, type CellRangeRef, type CellStyleRecord } from '@bilig/protocol'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { getWorkbookShortcutLabel } from '../shortcut-registry.js'
import { WorkbookToolbar } from '../workbook-toolbar.js'
import { deriveWorkbookStatusPresentation, useWorkbookToolbar } from '../use-workbook-toolbar.js'

function ToolbarHookHarness(props: {
  readonly invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  readonly selectionRangeRef: MutableRefObject<CellRangeRef>
  readonly selectedStyle?: CellStyleRecord | undefined
  readonly writesAllowed?: boolean | undefined
}) {
  const { ribbon } = useWorkbookToolbar({
    canHideCurrentColumn: false,
    canHideCurrentRow: false,
    canRedo: false,
    canUndo: false,
    canUnhideCurrentColumn: false,
    canUnhideCurrentRow: false,
    connectionStateName: 'connected',
    currentFillColor: '#ffffff',
    currentNumberFormatKind: 'general',
    currentTextColor: '#111827',
    horizontalAlignment: null,
    invokeMutation: props.invokeMutation,
    localPersistenceMode: 'persistent',
    onApplyBorderPreset: () => {},
    onClearStyle: () => {},
    onFillColorReset: () => {},
    onFillColorSelect: () => {},
    onFontSizeChange: () => {},
    onHideCurrentColumn: () => {},
    onHideCurrentRow: () => {},
    onHorizontalAlignmentChange: () => {},
    onNumberFormatChange: () => {},
    onRedo: () => {},
    onTextColorReset: () => {},
    onTextColorSelect: () => {},
    onToggleBold: () => {},
    onToggleItalic: () => {},
    onToggleUnderline: () => {},
    onToggleWrap: () => {},
    onUndo: () => {},
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

  it('derives clear visible save states for saved, saving, local-only, offline, and sync issues', () => {
    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connected',
        runtimeReady: true,
        localPersistenceMode: 'persistent',
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
        localPersistenceMode: 'persistent',
        pendingMutationSummary: { activeCount: 2, failedCount: 0 },
        remoteSyncAvailable: true,
        zeroConfigured: true,
        zeroHealthReady: true,
        writesAllowed: true,
      }),
    ).toMatchObject({
      syncLabel: 'Saving…',
      tone: 'progress',
    })

    expect(
      deriveWorkbookStatusPresentation({
        connectionStateName: 'connecting',
        runtimeReady: true,
        localPersistenceMode: 'persistent',
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
        localPersistenceMode: 'persistent',
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
        localPersistenceMode: 'persistent',
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
        connectionStateName: 'disconnected',
        runtimeReady: true,
        localPersistenceMode: 'persistent',
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
        localPersistenceMode: 'persistent',
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
    expect(fontSizeTrigger?.className).toContain('h-8')
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
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'i', metaKey: true }))
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Underline']")?.className).toContain('bg-[var(--wb-accent-soft)]')

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

    for (const resolveMutation of pendingMutationResolvers) {
      resolveMutation()
    }
    await act(async () => {
      await Promise.all(pendingMutations)
    })

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
})
