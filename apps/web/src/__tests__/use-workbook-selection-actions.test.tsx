// @vitest-environment jsdom
import { act, createElement, useEffect } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, test, vi } from 'vitest'
import type { GridSelectionSnapshot } from '@bilig/grid'
import { useWorkbookSelectionActions } from '../use-workbook-selection-actions.js'

interface HarnessProps {
  capture: (actions: ReturnType<typeof useWorkbookSelectionActions>) => void
  invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  supersedeOptimisticCellSeedsForRange?: Parameters<typeof useWorkbookSelectionActions>[0]['supersedeOptimisticCellSeedsForRange']
  replaceOptimisticCellSeed?: Parameters<typeof useWorkbookSelectionActions>[0]['replaceOptimisticCellSeed']
}

function SelectionActionsHarness({
  capture,
  invokeMutation,
  supersedeOptimisticCellSeedsForRange,
  replaceOptimisticCellSeed,
}: HarnessProps) {
  const actions = useWorkbookSelectionActions({
    writesAllowed: true,
    selectionRangeRef: {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    },
    selectionRef: {
      current: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
    },
    editorTargetRef: {
      current: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
    },
    editorValueRef: { current: '' },
    editingModeRef: { current: 'idle' },
    invokeMutation,
    applyParsedInput: vi.fn(),
    supersedeOptimisticCellSeedsForRange,
    replaceOptimisticCellSeed,
    resetEditorConflictTracking: vi.fn(),
    reportRuntimeError: vi.fn(),
    setEditorValue: vi.fn(),
    setEditingMode: vi.fn(),
    setEditorSelectionBehavior: vi.fn(),
  })

  useEffect(() => {
    capture(actions)
  }, [actions, capture])

  return createElement('div')
}

describe('useWorkbookSelectionActions', () => {
  test('clears an explicit grid selection snapshot instead of stale selection refs', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => undefined)
    let capturedActions: ReturnType<typeof useWorkbookSelectionActions> | null = null
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(SelectionActionsHarness, {
          capture: (actions: ReturnType<typeof useWorkbookSelectionActions>) => {
            capturedActions = actions
          },
          invokeMutation,
        }),
      )
    })

    const visibleSelection: GridSelectionSnapshot = {
      sheetName: 'Sheet1',
      address: 'C3',
      kind: 'range',
      range: {
        startAddress: 'C3',
        endAddress: 'D4',
      },
    }

    await act(async () => {
      capturedActions?.clearSelectedCell(visibleSelection)
      await Promise.resolve()
    })

    expect(invokeMutation).toHaveBeenCalledWith('clearRange', {
      sheetName: 'Sheet1',
      startAddress: 'C3',
      endAddress: 'D4',
    })
    expect(invokeMutation).not.toHaveBeenCalledWith('clearRange', {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A1',
    })

    await act(async () => {
      root.unmount()
    })
  })

  test('supersedes optimistic formula-bar seeds when clearing a selected range', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => undefined)
    const supersedeOptimisticCellSeedsForRange = vi.fn(() => null)
    const replaceOptimisticCellSeed = vi.fn(() => null)
    let capturedActions: ReturnType<typeof useWorkbookSelectionActions> | null = null
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(SelectionActionsHarness, {
          capture: (actions: ReturnType<typeof useWorkbookSelectionActions>) => {
            capturedActions = actions
          },
          invokeMutation,
          supersedeOptimisticCellSeedsForRange,
          replaceOptimisticCellSeed,
        }),
      )
    })

    const visibleSelection: GridSelectionSnapshot = {
      sheetName: 'Sheet1',
      address: 'B2',
      kind: 'cell',
      range: {
        startAddress: 'B2',
        endAddress: 'B2',
      },
    }

    await act(async () => {
      capturedActions?.clearSelectedCell(visibleSelection)
      await Promise.resolve()
    })

    expect(supersedeOptimisticCellSeedsForRange).toHaveBeenCalledWith({
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'B2',
    })
    expect(replaceOptimisticCellSeed).toHaveBeenCalledWith('Sheet1', 'B2', '')

    await act(async () => {
      root.unmount()
    })
  })

  test('restores superseded optimistic seeds when clear mutation fails', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {
      throw new Error('clear failed')
    })
    const rollbackOptimisticSeeds = vi.fn()
    const rollbackSelectedCellSeed = vi.fn()
    const supersedeOptimisticCellSeedsForRange = vi.fn(() => rollbackOptimisticSeeds)
    const replaceOptimisticCellSeed = vi.fn(() => rollbackSelectedCellSeed)
    let capturedActions: ReturnType<typeof useWorkbookSelectionActions> | null = null
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(SelectionActionsHarness, {
          capture: (actions: ReturnType<typeof useWorkbookSelectionActions>) => {
            capturedActions = actions
          },
          invokeMutation,
          supersedeOptimisticCellSeedsForRange,
          replaceOptimisticCellSeed,
        }),
      )
    })

    await act(async () => {
      capturedActions?.clearSelectedCell()
      await Promise.resolve()
    })

    expect(rollbackOptimisticSeeds).toHaveBeenCalledTimes(1)
    expect(rollbackSelectedCellSeed).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })
})
