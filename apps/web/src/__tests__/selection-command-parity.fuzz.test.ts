// @vitest-environment jsdom
import { act, createElement, type MutableRefObject } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi, afterEach } from 'vitest'
import * as fc from 'fast-check'
import { ValueTag, type CellRangeRef } from '@bilig/protocol'
import { runProperty } from '@bilig/test-fuzz'
import { useWorkbookToolbar } from '../use-workbook-toolbar.js'

function ToolbarHarness(props: {
  readonly invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  readonly selectionRangeRef: MutableRefObject<CellRangeRef>
}) {
  const { ribbon } = useWorkbookToolbar({
    connectionStateName: 'connected',
    runtimeReady: true,
    localPersistenceMode: 'persistent',
    remoteSyncAvailable: true,
    zeroConfigured: true,
    zeroHealthReady: true,
    canUndo: false,
    canRedo: false,
    onUndo: () => {},
    onRedo: () => {},
    canHideCurrentRow: false,
    canHideCurrentColumn: false,
    canUnhideCurrentRow: false,
    canUnhideCurrentColumn: false,
    onHideCurrentRow: () => {},
    onHideCurrentColumn: () => {},
    onUnhideCurrentRow: () => {},
    onUnhideCurrentColumn: () => {},
    invokeMutation: props.invokeMutation,
    selectionRangeRef: props.selectionRangeRef,
    selectedCell: {
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
    },
    selectedStyle: undefined,
    writesAllowed: true,
    trailingContent: null,
  })
  return createElement('div', null, ribbon)
}

afterEach(() => {
  document.body.innerHTML = ''
})

describe('selection command parity fuzz', () => {
  it('dispatches formatting mutations against the current live range without rerendering', async () => {
    await runProperty({
      suite: 'web/selection-command/live-range-parity',
      arbitrary: fc.array(selectionRangeArbitrary, { minLength: 2, maxLength: 12 }),
      predicate: async (ranges) => {
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

        try {
          await act(async () => {
            root.render(
              createElement(ToolbarHarness, {
                invokeMutation,
                selectionRangeRef,
              }),
            )
          })

          await ranges.reduce<Promise<void>>(async (previous, range) => {
            await previous
            selectionRangeRef.current = range
            await act(async () => {
              host.querySelector("[aria-label='Bold']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
            })
          }, Promise.resolve())

          const dispatchedRanges = invokeMutation.mock.calls.map((call) => call[1])
          expect(dispatchedRanges).toEqual(ranges)
        } finally {
          await act(async () => {
            root.unmount()
          })
        }
      },
    })
  })
})

const selectionRangeArbitrary = fc.constantFrom<CellRangeRef>(
  { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
  { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'D5' },
  { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'C9' },
  { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'F4' },
  { sheetName: 'Sheet1', startAddress: 'E1', endAddress: 'G3' },
)
