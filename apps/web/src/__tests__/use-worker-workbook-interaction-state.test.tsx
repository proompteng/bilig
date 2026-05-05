// @vitest-environment jsdom
import { act, createElement, useEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridSelectionSnapshot } from '@bilig/grid'
import type { WorkerHandle, WorkerRuntimeSelection } from '../runtime-session.js'
import { useWorkerWorkbookInteractionState } from '../use-worker-workbook-interaction-state.js'

function InteractionHarness(props: {
  documentId: string
  selection: WorkerRuntimeSelection
  selectedCell: CellSnapshot
  workerHandle: WorkerHandle | null
  invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  sendSelectionChanged: (selection: WorkerRuntimeSelection) => void
  onSelectionSheetChanged?: (nextSelection: WorkerRuntimeSelection, previousSelection: WorkerRuntimeSelection) => void
  capture: (value: ReturnType<typeof useWorkerWorkbookInteractionState>) => void
}) {
  const state = useWorkerWorkbookInteractionState({
    documentId: props.documentId,
    selection: props.selection,
    selectedCell: props.selectedCell,
    workerHandle: props.workerHandle,
    workerHandleRef: { current: props.workerHandle },
    writesAllowed: true,
    invokeMutation: props.invokeMutation,
    perfSession: {
      scope: 'test',
      markShellMounted() {},
      noteBootstrapResult() {},
      markFirstAuthoritativePatchVisible() {},
      markFirstReconcileStarted() {},
      markFirstReconcileSettled() {},
      markFirstSelectionVisible() {},
      markFirstLocalEditApplied() {},
      markFirstPasteApplied() {},
    },
    reportRuntimeError: vi.fn(),
    sendSelectionChanged: props.sendSelectionChanged,
    onSelectionSheetChanged: props.onSelectionSheetChanged,
  })

  useEffect(() => {
    props.capture(state)
  }, [props, state])

  return createElement('div')
}

function mountHarness(): {
  root: Root
  render: (props: Parameters<typeof InteractionHarness>[0]) => Promise<void>
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  const render = async (props: Parameters<typeof InteractionHarness>[0]) => {
    await act(async () => {
      root.render(createElement(InteractionHarness, props))
    })
  }
  return { root, render }
}

describe('useWorkerWorkbookInteractionState', () => {
  afterEach(() => {
    document.body.innerHTML = ''
    vi.clearAllMocks()
  })

  it('tracks external selection changes and resets the viewport when switching sheets', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const sendSelectionChanged = vi.fn()
    const onSelectionSheetChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell: stringCell('Sheet1', 'A1', 'one'),
      workerHandle: { viewportStore: createViewportStoreStub('Sheet1', 'A1', stringCell('Sheet1', 'A1', 'one')) },
      invokeMutation: vi.fn(async () => undefined),
      sendSelectionChanged,
      onSelectionSheetChanged,
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.selectAddress('Sheet2', 'B3')
    })

    expect(sendSelectionChanged).toHaveBeenCalledWith({ sheetName: 'Sheet2', address: 'B3' })
    expect(onSelectionSheetChanged).toHaveBeenCalledWith({ sheetName: 'Sheet2', address: 'B3' }, { sheetName: 'Sheet1', address: 'A1' })

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('commits editor changes through the extracted controller', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', 'before')
    const workerHandle = { viewportStore: createViewportStoreStub('Sheet1', 'A1', selectedCell) }
    const invokeMutation = vi.fn(async () => undefined)
    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell,
      workerHandle,
      invokeMutation,
      sendSelectionChanged,
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.beginEditing()
      captured?.handleEditorChange('after')
      captured?.commitEditor()
      await Promise.resolve()
    })

    expect(invokeMutation).toHaveBeenCalledWith('setCellValue', 'Sheet1', 'A1', 'after')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('accepts user selection after the grid acknowledges an external selection', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell: stringCell('Sheet1', 'A1', 'one'),
      workerHandle: { viewportStore: createViewportStoreStub('Sheet1', 'A1', stringCell('Sheet1', 'A1', 'one')) },
      invokeMutation: vi.fn(async () => undefined),
      sendSelectionChanged,
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.selectAddress('Sheet1', 'B2')
    })
    expect(captured.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'B2' })

    await act(async () => {
      captured?.handleSelectionChange(singleCellSnapshot('Sheet1', 'C3'))
    })
    expect(captured.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'B2' })

    await act(async () => {
      captured?.acknowledgeExternalSelectionSync(singleCellSnapshot('Sheet1', 'B2'))
      captured?.handleSelectionChange(singleCellSnapshot('Sheet1', 'C3'))
    })
    expect(captured.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'C3' })

    await act(async () => {
      harness.root.unmount()
    })
  })
})

function createViewportStoreStub(sheetName: string, address: string, cell: CellSnapshot) {
  return {
    getCell(targetSheetName: string, targetAddress: string) {
      if (targetSheetName === sheetName && targetAddress === address) {
        return cell
      }
      return stringCell(targetSheetName, targetAddress, '')
    },
    setCellSnapshot: vi.fn(),
  }
}

function stringCell(sheetName: string, address: string, value: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: {
      tag: ValueTag.String,
      value,
    },
    input: value,
    flags: 0,
    version: 1,
  }
}

function singleCellSnapshot(sheetName: string, address: string): GridSelectionSnapshot {
  return {
    sheetName,
    address,
    kind: 'cell',
    range: {
      sheetName,
      startAddress: address,
      endAddress: address,
    },
  }
}
