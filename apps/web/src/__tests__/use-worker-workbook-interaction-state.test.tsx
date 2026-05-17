// @vitest-environment jsdom
import { act, createElement, useEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag, type CellSnapshot } from '@bilig/protocol'
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

  it('commits a cleared formula bar draft before applying a click-away selection', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', 'stale')
    const workerHandle = {
      viewportStore: createViewportStoreMapStub([selectedCell, stringCell('Sheet1', 'B1', 'next')]),
    }
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
      captured?.handleEditorChange('')
      captured?.handleSelectionChange(singleCellSnapshot('Sheet1', 'B1'))
      await Promise.resolve()
    })

    expect(invokeMutation).toHaveBeenCalledWith('clearCell', 'Sheet1', 'A1')
    expect(workerHandle.viewportStore.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })
    expect(captured?.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'B1' })
    expect(captured?.visibleEditorValue).toBe('next')
    expect(sendSelectionChanged).toHaveBeenCalledWith({ sheetName: 'Sheet1', address: 'B1' })

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('commits a formula bar value override even before editing state catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', '')
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
      captured?.commitEditor(undefined, '=A1="HELLO"')
      await Promise.resolve()
    })

    expect(invokeMutation).toHaveBeenCalledWith('setCellFormula', 'Sheet1', 'A1', 'A1="HELLO"')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('commits an empty formula bar override even when the edit base is stale', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', '')
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
      captured?.beginEditing('', 'select-all', 'formula')
      workerHandle.viewportStore.setCellSnapshot(stringCell('Sheet1', 'A1', 'stale value'))
      captured?.commitEditor(undefined, '')
      await Promise.resolve()
    })

    expect(invokeMutation).toHaveBeenCalledWith('clearCell', 'Sheet1', 'A1')
    expect(workerHandle.viewportStore.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('keeps invalid formula commits visible until selected-cell state catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', '')
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
      captured?.commitEditor(undefined, '=1+')
      await Promise.resolve()
    })

    expect(workerHandle.viewportStore.setCellSnapshot).toHaveBeenCalledWith(
      expect.objectContaining({
        input: '=1+',
        value: { tag: ValueTag.Error, code: ErrorCode.Value },
      }),
    )
    expect(captured?.visibleEditorValue).toBe('#VALUE!')
    expect(invokeMutation).toHaveBeenCalledWith('setCellFormula', 'Sheet1', 'A1', '1+')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('keeps optimistic formula results visible until selected-cell state catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A2', '')
    const workerHandle = {
      viewportStore: createViewportStoreMapStub([stringCell('Sheet1', 'A1', 'HELLO'), selectedCell]),
    }
    const invokeMutation = vi.fn(async () => undefined)
    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A2' },
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
      captured?.commitEditor(undefined, '=A1="HELLO"')
      await Promise.resolve()
    })

    expect(workerHandle.viewportStore.getCell('Sheet1', 'A2').value).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(captured?.visibleResolvedValue).toBe('TRUE')
    expect(invokeMutation).toHaveBeenCalledWith('setCellFormula', 'Sheet1', 'A2', 'A1="HELLO"')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('keeps detached optimistic value readback visible before the viewport store is ready', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', '')
    const invokeMutation = vi.fn(async () => undefined)
    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell,
      workerHandle: null,
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
      captured?.commitEditor(undefined, '12')
      await Promise.resolve()
      await Promise.resolve()
    })

    expect(captured?.visibleEditorValue).toBe('12')
    expect(captured?.visibleResolvedValue).toBe('12')
    expect(invokeMutation).toHaveBeenCalledWith('setCellValue', 'Sheet1', 'A1', 12)

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('keeps local editor seeds until live viewport readback catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const initialCell = stringCell('Sheet1', 'A1', '')
    const authoritativeCell = stringCell('Sheet1', 'A1', '12')
    let viewportCell = initialCell
    const workerHandle = {
      viewportStore: {
        getCell: () => viewportCell,
        setCellSnapshot: vi.fn(),
      },
    }
    const invokeMutation = vi.fn(async () => undefined)
    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null
    const baseProps = {
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      workerHandle,
      invokeMutation,
      sendSelectionChanged,
      capture: (value: ReturnType<typeof useWorkerWorkbookInteractionState>) => {
        captured = value
      },
    }

    await harness.render({
      ...baseProps,
      selectedCell: initialCell,
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.commitEditor(undefined, '12')
      await Promise.resolve()
    })

    await harness.render({
      ...baseProps,
      selectedCell: authoritativeCell,
    })

    expect(captured?.getCellEditorSeed('Sheet1', 'A1')).toBe('12')

    viewportCell = authoritativeCell
    await harness.render({
      ...baseProps,
      selectedCell: authoritativeCell,
    })

    expect(captured?.getCellEditorSeed('Sheet1', 'A1')).toBeUndefined()
    expect(captured?.visibleResolvedValue).toBe('12')

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('can supersede and restore sheet-wide optimistic seeds around structural edits', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', '')
    const workerHandle = null
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
      captured?.commitEditor(undefined, '12')
      await Promise.resolve()
      await Promise.resolve()
    })

    expect(captured?.getCellEditorSeed('Sheet1', 'A1')).toBe('12')

    let rollback: (() => void) | null | undefined
    await act(async () => {
      rollback = captured?.supersedeOptimisticCellSeedsForSheet('Sheet1')
    })

    expect(captured?.getCellEditorSeed('Sheet1', 'A1')).toBeUndefined()

    await act(async () => {
      rollback?.()
    })

    expect(captured?.getCellEditorSeed('Sheet1', 'A1')).toBe('12')

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

  it('keeps an external selection visible while a stale authoritative selection update catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    const workerHandle = {
      viewportStore: createViewportStoreMapStub([
        stringCell('Sheet1', 'A1', 'one'),
        stringCell('Sheet1', 'B2', 'two'),
        stringCell('Sheet1', 'C3', 'three'),
      ]),
    }
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null
    const baseProps = {
      documentId: 'doc-1',
      selectedCell: stringCell('Sheet1', 'A1', 'one'),
      workerHandle,
      invokeMutation: vi.fn(async () => undefined),
      sendSelectionChanged,
      capture: (value: ReturnType<typeof useWorkerWorkbookInteractionState>) => {
        captured = value
      },
    }

    await harness.render({
      ...baseProps,
      selection: { sheetName: 'Sheet1', address: 'A1' },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.selectAddress('Sheet1', 'B2')
    })
    expect(captured.visibleSelection).toEqual({ sheetName: 'Sheet1', address: 'B2' })

    await harness.render({
      ...baseProps,
      selection: { sheetName: 'Sheet1', address: 'C3' },
    })
    expect(captured.visibleSelection).toEqual({ sheetName: 'Sheet1', address: 'B2' })
    expect(captured.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'B2' })

    await harness.render({
      ...baseProps,
      selection: { sheetName: 'Sheet1', address: 'B2' },
    })
    expect(captured.visibleSelection).toEqual({ sheetName: 'Sheet1', address: 'B2' })

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('updates visible selection and editor readback before authoritative selection catches up', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const sendSelectionChanged = vi.fn()
    const selectedCell = stringCell('Sheet1', 'A1', 'one')
    const workerHandle = {
      viewportStore: createViewportStoreMapStub([selectedCell, stringCell('Sheet1', 'B2', 'two')]),
    }
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell,
      workerHandle,
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
      captured?.handleSelectionChange(singleCellSnapshot('Sheet1', 'B2'))
    })

    expect(captured?.visibleSelection).toEqual({ sheetName: 'Sheet1', address: 'B2' })
    expect(captured?.visibleSelectedCell).toMatchObject({ sheetName: 'Sheet1', address: 'B2', input: 'two' })
    expect(captured?.visibleEditorValue).toBe('two')
    expect(captured?.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'B2' })
    expect(sendSelectionChanged).toHaveBeenCalledWith({ sheetName: 'Sheet1', address: 'B2' })

    await act(async () => {
      harness.root.unmount()
    })
  })

  it('updates visible readback during idle selection navigation without starting edit conflict tracking', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const sendSelectionChanged = vi.fn()
    const initialCell = stringCell('Sheet1', 'A1', 'one')
    const getCell = vi.fn((sheetName: string, address: string) => stringCell(sheetName, address, address))
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell: initialCell,
      workerHandle: {
        viewportStore: {
          getCell,
          setCellSnapshot: vi.fn(),
        },
      },
      invokeMutation: vi.fn(async () => undefined),
      sendSelectionChanged,
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    getCell.mockClear()
    await act(async () => {
      captured?.handleSelectionChange(singleCellSnapshot('Sheet1', 'C3'))
    })

    expect(captured.selectionRef.current).toEqual({ sheetName: 'Sheet1', address: 'C3' })
    expect(sendSelectionChanged).toHaveBeenCalledWith({ sheetName: 'Sheet1', address: 'C3' })
    expect(captured.visibleEditorValue).toBe('C3')
    expect(getCell).toHaveBeenCalledWith('Sheet1', 'C3')
    const visibleReadbackCallCount = getCell.mock.calls.length

    await act(async () => {
      captured?.beginEditing()
    })

    expect(getCell).toHaveBeenCalledWith('Sheet1', 'C3')
    expect(getCell.mock.calls.length).toBeGreaterThan(visibleReadbackCallCount)

    await act(async () => {
      harness.root.unmount()
    })
  })
})

function createViewportStoreStub(sheetName: string, address: string, cell: CellSnapshot) {
  let activeCell = cell
  return {
    getCell(targetSheetName: string, targetAddress: string) {
      if (targetSheetName === sheetName && targetAddress === address) {
        return activeCell
      }
      return stringCell(targetSheetName, targetAddress, '')
    },
    setCellSnapshot: vi.fn((snapshot: CellSnapshot) => {
      if (snapshot.sheetName === sheetName && snapshot.address === address) {
        activeCell = snapshot
      }
    }),
  }
}

function createViewportStoreMapStub(cells: readonly CellSnapshot[]) {
  const cellMap = new Map(cells.map((cell) => [`${cell.sheetName}!${cell.address}`, cell] as const))
  return {
    getCell(targetSheetName: string, targetAddress: string) {
      return cellMap.get(`${targetSheetName}!${targetAddress}`) ?? stringCell(targetSheetName, targetAddress, '')
    },
    setCellSnapshot: vi.fn((snapshot: CellSnapshot) => {
      cellMap.set(`${snapshot.sheetName}!${snapshot.address}`, snapshot)
    }),
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
