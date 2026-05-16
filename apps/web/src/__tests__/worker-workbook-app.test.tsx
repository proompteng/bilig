// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import type { Action, ToastT, ToastToDismiss } from 'sonner'
import { toast } from 'sonner'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkerWorkbookApp } from '../WorkerWorkbookApp.js'

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

function findActiveToast(id: string): ToastT | null {
  return toast.getToasts().find((entry: ToastT | ToastToDismiss): entry is ToastT => !('dismiss' in entry) && entry.id === id) ?? null
}

function getToastAction(activeToast: ToastT | null): Action {
  if (!activeToast || !activeToast.action || typeof activeToast.action !== 'object' || !('onClick' in activeToast.action)) {
    throw new Error('Expected toast action object')
  }
  return activeToast.action
}

const { latestWorkbookViewProps, useWorkerWorkbookAppState } = vi.hoisted(() => ({
  latestWorkbookViewProps: { current: null as Record<string, unknown> | null },
  useWorkerWorkbookAppState: vi.fn(),
}))

vi.mock('@bilig/grid', () => ({
  WorkbookView: (props: Record<string, unknown>) => {
    latestWorkbookViewProps.current = props
    return <div data-testid="workbook-view" />
  },
}))

vi.mock('../use-worker-workbook-app-state.js', () => ({
  useWorkerWorkbookAppState,
}))

vi.mock('../use-workbook-import-pane.js', () => ({
  useWorkbookImportPane: () => ({
    clearImportError: vi.fn(),
    importError: null,
    importPanel: null,
    importToggle: null,
  }),
}))

vi.mock('../use-workbook-shortcut-dialog.js', () => ({
  useWorkbookShortcutDialog: () => ({
    shortcutDialog: null,
    shortcutHelpButton: null,
  }),
}))

function createCellSelectionSnapshot(sheetName: string, address: string): Record<string, unknown> {
  return {
    sheetName,
    address,
    kind: 'cell',
    range: {
      startAddress: address,
      endAddress: address,
    },
  }
}

function createReadyWorkbookAppState(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    agentError: null,
    approvePersistenceTransfer: vi.fn(),
    autofitColumn: vi.fn(),
    beginEditing: vi.fn(),
    canRedo: false,
    canUndo: false,
    cancelEditor: vi.fn(),
    changeCount: 0,
    changesPanel: null,
    clearAgentError: vi.fn(),
    clearRuntimeError: vi.fn(),
    clearSelectedCell: vi.fn(),
    columnWidths: {},
    commitEditor: vi.fn(),
    copySelectionRange: vi.fn(),
    createSheet: vi.fn(),
    definedNames: [],
    deleteSheet: vi.fn(),
    dismissPersistenceTransferRequest: vi.fn(),
    editorConflictBanner: null,
    editorSelectionBehavior: 'select-all',
    failedPendingMutation: null,
    fillSelectionRange: vi.fn(),
    freezeCols: 0,
    freezeRows: 0,
    getCellEditorSeed: vi.fn(),
    acknowledgeExternalSelectionSync: vi.fn(),
    handleEditorChange: vi.fn(),
    handleSelectionChange: vi.fn(),
    handleVisibleViewportChange: vi.fn(),
    hiddenColumns: {},
    hiddenRows: {},
    invokeColumnVisibilityMutation: vi.fn(),
    invokeColumnWidthMutation: vi.fn(),
    invokeDeleteColumnsMutation: vi.fn(),
    invokeDeleteRowsMutation: vi.fn(),
    invokeInsertColumnsMutation: vi.fn(),
    invokeInsertRowsMutation: vi.fn(),
    invokeRowHeightMutation: vi.fn(),
    invokeRowVisibilityMutation: vi.fn(),
    invokeSetFreezePaneMutation: vi.fn(),
    isEditing: false,
    isEditingCell: false,
    localPersistenceMode: 'ephemeral',
    moveSelectionRange: vi.fn(),
    pasteIntoSelection: vi.fn(),
    pendingTransferRequest: null,
    previewRanges: [],
    redoLatestChange: vi.fn(),
    remoteSyncAvailable: true,
    renameSheet: vi.fn(),
    reportRuntimeError: vi.fn(),
    requestPersistenceTransfer: vi.fn(),
    resolvedValue: '',
    retryFailedPendingMutation: vi.fn(),
    ribbon: null,
    rowHeights: {},
    runtimeError: null,
    runtimeReady: true,
    runtimeSyncState: 'local-only',
    selection: { sheetName: 'Sheet1', address: 'A1' },
    selectionSnapshot: createCellSelectionSnapshot('Sheet1', 'A1'),
    selectAddress: vi.fn(),
    selectSelectionSnapshot: vi.fn(),
    selectedCell: { sheetName: 'Sheet1', address: 'A1' },
    setSidePanelWidth: vi.fn(),
    sheetIdsByName: { 'Prepaid Template': 1, Sheet1: 2 },
    sheetNames: ['Prepaid Template', 'Sheet1'],
    sheetOrdinalsByName: { 'Prepaid Template': 0, Sheet1: 1 },
    sidePanel: null,
    sidePanelId: undefined,
    sidePanelWidth: undefined,
    statusModeLabel: 'Live',
    toggleBooleanCell: vi.fn(),
    toolbarTrailingContent: null,
    transferRequested: false,
    undoLatestChange: vi.fn(),
    visibleEditorValue: '',
    visibleSelectedCell: { sheetName: 'Sheet1', address: 'A1' },
    visibleSelection: { sheetName: 'Sheet1', address: 'A1' },
    workbookReady: true,
    workerHandle: {
      viewportStore: {},
    },
    writesAllowed: true,
    zeroConfigured: true,
    ...overrides,
  }
}

describe('WorkerWorkbookApp', () => {
  afterEach(() => {
    toast.dismiss()
    vi.clearAllMocks()
    latestWorkbookViewProps.current = null
    document.body.innerHTML = ''
    window.history.replaceState({}, '', '/')
  })

  it('renders a retry toast for failed pending mutations', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const retryFailedPendingMutation = vi.fn()
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: {
        id: 'pending-1',
        method: 'setCellValue',
        failureMessage: 'mutation rejected by server',
        attemptCount: 1,
      },
      approvePersistenceTransfer: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      pendingTransferRequest: null,
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation,
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: false,
      localPersistenceMode: 'ephemeral',
      statusModeLabel: 'Live',
      transferRequested: false,
      workbookReady: false,
      workerHandle: null,
      zeroConfigured: true,
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })
    await flushToasts()

    const pendingToast = findActiveToast('pending-mutation-pending-1')
    expect(pendingToast?.title).toBe('A local change could not be synced. mutation rejected by server')
    const retryAction = getToastAction(pendingToast)

    await act(async () => {
      Reflect.apply(retryAction.onClick, undefined, [new MouseEvent('click')])
    })

    expect(retryFailedPendingMutation).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('passes one authoritative selection-change callback into the workbook view', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer: vi.fn(),
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      pendingTransferRequest: null,
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      workbookReady: true,
      zeroConfigured: true,
      localPersistenceMode: 'ephemeral',
      statusModeLabel: 'Live',
      transferRequested: false,
      workerHandle: {
        viewportStore: {},
      },
      handleSelectionChange: vi.fn(),
      selection: { sheetName: 'Sheet1', address: 'B18' },
      selectedCell: { sheetName: 'Sheet1', address: 'B18' },
      sheetNames: ['Sheet1'],
      previewRanges: [],
      resolvedValue: '',
      sidePanel: null,
      sidePanelId: undefined,
      sidePanelWidth: undefined,
      setSidePanelWidth: vi.fn(),
      commitEditor: vi.fn(),
      copySelectionRange: vi.fn(),
      createSheet: vi.fn(),
      fillSelectionRange: vi.fn(),
      handleEditorChange: vi.fn(),
      handleVisibleViewportChange: vi.fn(),
      invokeColumnVisibilityMutation: vi.fn(),
      invokeDeleteColumnsMutation: vi.fn(),
      invokeDeleteRowsMutation: vi.fn(),
      invokeInsertColumnsMutation: vi.fn(),
      invokeInsertRowsMutation: vi.fn(),
      invokeRowVisibilityMutation: vi.fn(),
      invokeSetFreezePaneMutation: vi.fn(),
      isEditing: false,
      isEditingCell: false,
      moveSelectionRange: vi.fn(),
      pasteIntoSelection: vi.fn(),
      renameSheet: vi.fn(),
      reportRuntimeError: vi.fn(),
      runtimeErrorBanner: null,
      selectAddress: vi.fn(),
      approveBundle: vi.fn(),
      ribbon: null,
      dismissPersistenceTransferRequestBanner: null,
      subscribeViewport: vi.fn(),
      columnWidths: {},
      hiddenColumns: {},
      hiddenRows: {},
      rowHeights: {},
      freezeRows: 0,
      freezeCols: 0,
      toggleBooleanCell: vi.fn(),
      visibleEditorValue: '',
      importPanel: null,
      importToggle: null,
      clearImportError: vi.fn(),
      pendingCommandCount: 0,
      canUndo: false,
      canRedo: false,
      changeCount: 0,
      changesPanel: null,
      redoLatestChange: vi.fn(),
      undoLatestChange: vi.fn(),
      startNewThread: vi.fn(),
      requestPersistenceTransferBanner: null,
      localPersistenceBanner: null,
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })

    expect(typeof latestWorkbookViewProps.current?.['onSelectionChange']).toBe('function')
    expect(latestWorkbookViewProps.current?.['onSelectionRangeChange']).toBeUndefined()

    await act(async () => {
      root.unmount()
    })
  })

  it('routes typed range targets from the name box into the authoritative selection snapshot path', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectSelectionSnapshot = vi.fn()

    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer: vi.fn(),
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      pendingTransferRequest: null,
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      workbookReady: true,
      zeroConfigured: true,
      localPersistenceMode: 'ephemeral',
      statusModeLabel: 'Live',
      transferRequested: false,
      workerHandle: {
        viewportStore: {},
      },
      handleSelectionChange: vi.fn(),
      selection: { sheetName: 'Sheet1', address: 'B2' },
      selectionSnapshot: {
        sheetName: 'Sheet1',
        address: 'B2',
        kind: 'cell',
        range: {
          startAddress: 'B2',
          endAddress: 'B2',
        },
      },
      selectedCell: { sheetName: 'Sheet1', address: 'B2' },
      sheetNames: ['Sheet1'],
      previewRanges: [],
      resolvedValue: '',
      sidePanel: null,
      sidePanelId: undefined,
      sidePanelWidth: undefined,
      setSidePanelWidth: vi.fn(),
      commitEditor: vi.fn(),
      copySelectionRange: vi.fn(),
      createSheet: vi.fn(),
      fillSelectionRange: vi.fn(),
      handleEditorChange: vi.fn(),
      handleVisibleViewportChange: vi.fn(),
      invokeColumnVisibilityMutation: vi.fn(),
      invokeDeleteColumnsMutation: vi.fn(),
      invokeDeleteRowsMutation: vi.fn(),
      invokeInsertColumnsMutation: vi.fn(),
      invokeInsertRowsMutation: vi.fn(),
      invokeRowVisibilityMutation: vi.fn(),
      invokeSetFreezePaneMutation: vi.fn(),
      isEditing: false,
      isEditingCell: false,
      moveSelectionRange: vi.fn(),
      pasteIntoSelection: vi.fn(),
      renameSheet: vi.fn(),
      reportRuntimeError: vi.fn(),
      runtimeErrorBanner: null,
      selectAddress: vi.fn(),
      selectSelectionSnapshot,
      approveBundle: vi.fn(),
      ribbon: null,
      dismissPersistenceTransferRequestBanner: null,
      subscribeViewport: vi.fn(),
      columnWidths: {},
      hiddenColumns: {},
      hiddenRows: {},
      rowHeights: {},
      freezeRows: 0,
      freezeCols: 0,
      toggleBooleanCell: vi.fn(),
      visibleEditorValue: '',
      importPanel: null,
      importToggle: null,
      clearImportError: vi.fn(),
      pendingCommandCount: 0,
      canUndo: false,
      canRedo: false,
      changeCount: 0,
      changesPanel: null,
      redoLatestChange: vi.fn(),
      undoLatestChange: vi.fn(),
      startNewThread: vi.fn(),
      requestPersistenceTransferBanner: null,
      localPersistenceBanner: null,
      definedNames: [],
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })

    const onAddressCommit = latestWorkbookViewProps.current?.['onAddressCommit']
    expect(typeof onAddressCommit).toBe('function')
    if (typeof onAddressCommit !== 'function') {
      throw new Error('WorkbookView did not receive an address commit handler')
    }

    await act(async () => {
      onAddressCommit('B2:D8')
    })

    expect(selectSelectionSnapshot).toHaveBeenCalledWith({
      sheetName: 'Sheet1',
      address: 'B2',
      kind: 'range',
      range: {
        startAddress: 'B2',
        endAddress: 'D8',
      },
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('renders a typed missing-sheet state instead of an empty workbook grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectAddress = vi.fn()
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer: vi.fn(),
      autofitColumn: vi.fn(),
      beginEditing: vi.fn(),
      canRedo: false,
      canUndo: false,
      cancelEditor: vi.fn(),
      changeCount: 0,
      changesPanel: null,
      clearAgentError: vi.fn(),
      clearImportError: vi.fn(),
      clearRuntimeError: vi.fn(),
      clearSelectedCell: vi.fn(),
      columnWidths: {},
      commitEditor: vi.fn(),
      copySelectionRange: vi.fn(),
      createSheet: vi.fn(),
      definedNames: [],
      deleteSheet: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      editorConflictBanner: null,
      editorSelectionBehavior: 'select-all',
      failedPendingMutation: null,
      fillSelectionRange: vi.fn(),
      freezeCols: 0,
      freezeRows: 0,
      getCellEditorSeed: vi.fn(),
      handleEditorChange: vi.fn(),
      handleSelectionChange: vi.fn(),
      handleVisibleViewportChange: vi.fn(),
      hiddenColumns: {},
      hiddenRows: {},
      importError: null,
      importPanel: null,
      importToggle: null,
      invokeColumnVisibilityMutation: vi.fn(),
      invokeColumnWidthMutation: vi.fn(),
      invokeDeleteColumnsMutation: vi.fn(),
      invokeDeleteRowsMutation: vi.fn(),
      invokeInsertColumnsMutation: vi.fn(),
      invokeInsertRowsMutation: vi.fn(),
      invokeRowHeightMutation: vi.fn(),
      invokeRowVisibilityMutation: vi.fn(),
      invokeSetFreezePaneMutation: vi.fn(),
      isEditing: false,
      isEditingCell: false,
      localPersistenceMode: 'ephemeral',
      localPersistenceBanner: null,
      moveSelectionRange: vi.fn(),
      pasteIntoSelection: vi.fn(),
      pendingTransferRequest: null,
      previewRanges: [],
      redoLatestChange: vi.fn(),
      remoteSyncAvailable: true,
      renameSheet: vi.fn(),
      reportRuntimeError: vi.fn(),
      requestPersistenceTransfer: vi.fn(),
      requestPersistenceTransferBanner: null,
      resolvedValue: '',
      retryFailedPendingMutation: vi.fn(),
      ribbon: null,
      rowHeights: {},
      runtimeError: null,
      runtimeReady: true,
      selection: { sheetName: 'Missing Sheet', address: 'E46' },
      selectionSnapshot: {
        sheetName: 'Missing Sheet',
        address: 'E46',
        kind: 'cell',
        range: {
          startAddress: 'E46',
          endAddress: 'E46',
        },
      },
      selectAddress,
      selectSelectionSnapshot: vi.fn(),
      selectedCell: { sheetName: 'Missing Sheet', address: 'E46' },
      setSidePanelWidth: vi.fn(),
      sheetIdsByName: { 'Prepaid Template': 1, Sheet1: 2 },
      sheetNames: ['Prepaid Template', 'Sheet1'],
      sheetOrdinalsByName: { 'Prepaid Template': 0, Sheet1: 1 },
      sidePanel: null,
      sidePanelId: undefined,
      sidePanelWidth: undefined,
      statusModeLabel: 'Live',
      toggleBooleanCell: vi.fn(),
      toolbarTrailingContent: null,
      transferRequested: false,
      undoLatestChange: vi.fn(),
      visibleEditorValue: '',
      visibleSelectedCell: { sheetName: 'Missing Sheet', address: 'E46' },
      visibleSelection: { sheetName: 'Missing Sheet', address: 'E46' },
      workbookReady: true,
      workerHandle: {
        viewportStore: {},
      },
      writesAllowed: true,
      zeroConfigured: true,
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })

    expect(host.querySelector("[data-testid='missing-sheet-state']")?.textContent).toContain('Missing Sheet')
    expect(latestWorkbookViewProps.current).toBeNull()

    const prepaidButton = [...host.querySelectorAll('button')].find((button) => button.textContent === 'Open Prepaid Template')
    expect(prepaidButton).toBeInstanceOf(HTMLButtonElement)
    if (!(prepaidButton instanceof HTMLButtonElement)) {
      throw new Error('Expected missing-sheet recovery button')
    }

    await act(async () => {
      prepaidButton.click()
    })

    expect(selectAddress).toHaveBeenCalledWith('Prepaid Template', 'A1')

    await act(async () => {
      root.unmount()
    })
  })

  it('applies same-document URL sheet and cell changes to workbook selection', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.history.replaceState({}, '', '/?sheet=Prepaid+Template&cell=F16')

    const selectAddress = vi.fn()
    useWorkerWorkbookAppState.mockReturnValue(createReadyWorkbookAppState({ selectAddress }))

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })

    expect(selectAddress).toHaveBeenCalledWith('Prepaid Template', 'F16')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not replay stale URL selection after a local click selection update', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.history.replaceState({}, '', '/?sheet=Prepaid+Template&cell=D53')

    const selectAddress = vi.fn()
    const initialState = createReadyWorkbookAppState({
      selectAddress,
      selection: { sheetName: 'Prepaid Template', address: 'D53' },
      selectionSnapshot: createCellSelectionSnapshot('Prepaid Template', 'D53'),
      selectedCell: { sheetName: 'Prepaid Template', address: 'D53' },
      visibleSelectedCell: { sheetName: 'Prepaid Template', address: 'D53' },
      visibleSelection: { sheetName: 'Prepaid Template', address: 'D53' },
    })
    let currentState = initialState
    useWorkerWorkbookAppState.mockImplementation(() => currentState)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const appElement = (
      <WorkerWorkbookApp
        config={{
          currentUserId: 'guest:test',
          defaultDocumentId: 'doc-1',
          persistState: true,
          zeroCacheUrl: 'http://127.0.0.1:4848',
        }}
        connectionState={{ name: 'connected' }}
      />
    )

    await act(async () => {
      root.render(appElement)
    })

    expect(selectAddress).not.toHaveBeenCalled()
    selectAddress.mockClear()
    currentState = {
      ...initialState,
      selection: { sheetName: 'Prepaid Template', address: 'E54' },
      selectionSnapshot: createCellSelectionSnapshot('Prepaid Template', 'E54'),
      selectedCell: { sheetName: 'Prepaid Template', address: 'E54' },
      visibleSelectedCell: { sheetName: 'Prepaid Template', address: 'E54' },
      visibleSelection: { sheetName: 'Prepaid Template', address: 'E54' },
    }

    await act(async () => {
      root.render(appElement)
    })

    expect(selectAddress).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  it('returns row and column delete promises to the workbook view context-menu handlers', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const deleteRowsTask = Promise.resolve()
    const deleteColumnsTask = Promise.resolve()
    const invokeDeleteRowsMutation = vi.fn(() => deleteRowsTask)
    const invokeDeleteColumnsMutation = vi.fn(() => deleteColumnsTask)
    const invokeColumnWidthMutation = vi.fn(() => Promise.resolve())
    const invokeRowHeightMutation = vi.fn(() => Promise.resolve())

    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer: vi.fn(),
      autofitColumn: vi.fn(),
      beginEditing: vi.fn(),
      canRedo: false,
      canUndo: false,
      cancelEditor: vi.fn(),
      changeCount: 0,
      changesPanel: null,
      clearAgentError: vi.fn(),
      clearImportError: vi.fn(),
      clearRuntimeError: vi.fn(),
      clearSelectedCell: vi.fn(),
      columnWidths: {},
      commitEditor: vi.fn(),
      copySelectionRange: vi.fn(),
      createSheet: vi.fn(),
      definedNames: [],
      deleteSheet: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      editorConflictBanner: null,
      editorSelectionBehavior: 'select-all',
      failedPendingMutation: null,
      fillSelectionRange: vi.fn(),
      freezeCols: 0,
      freezeRows: 0,
      getCellEditorSeed: vi.fn(),
      handleEditorChange: vi.fn(),
      handleSelectionChange: vi.fn(),
      handleVisibleViewportChange: vi.fn(),
      hiddenColumns: {},
      hiddenRows: {},
      importError: null,
      importPanel: null,
      importToggle: null,
      invokeColumnVisibilityMutation: vi.fn(),
      invokeColumnWidthMutation,
      invokeDeleteColumnsMutation,
      invokeDeleteRowsMutation,
      invokeInsertColumnsMutation: vi.fn(),
      invokeInsertRowsMutation: vi.fn(),
      invokeRowHeightMutation,
      invokeRowVisibilityMutation: vi.fn(),
      invokeSetFreezePaneMutation: vi.fn(),
      isEditing: false,
      isEditingCell: false,
      localPersistenceMode: 'ephemeral',
      localPersistenceBanner: null,
      moveSelectionRange: vi.fn(),
      pasteIntoSelection: vi.fn(),
      pendingTransferRequest: null,
      previewRanges: [],
      redoLatestChange: vi.fn(),
      remoteSyncAvailable: true,
      renameSheet: vi.fn(),
      reportRuntimeError: vi.fn(),
      requestPersistenceTransfer: vi.fn(),
      requestPersistenceTransferBanner: null,
      resolvedValue: '',
      retryFailedPendingMutation: vi.fn(),
      ribbon: null,
      rowHeights: {},
      runtimeError: null,
      runtimeReady: true,
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectionSnapshot: {
        sheetName: 'Sheet1',
        address: 'C3',
        kind: 'cell',
        range: {
          startAddress: 'C3',
          endAddress: 'C3',
        },
      },
      selectAddress: vi.fn(),
      selectSelectionSnapshot: vi.fn(),
      selectedCell: { sheetName: 'Sheet1', address: 'A1' },
      setSidePanelWidth: vi.fn(),
      sheetIdsByName: { Sheet1: 1, Sheet2: 2 },
      sheetNames: ['Sheet1', 'Sheet2'],
      sheetOrdinalsByName: { Sheet1: 0, Sheet2: 1 },
      sidePanel: null,
      sidePanelId: undefined,
      sidePanelWidth: undefined,
      statusModeLabel: 'Live',
      toggleBooleanCell: vi.fn(),
      toolbarTrailingContent: null,
      transferRequested: false,
      undoLatestChange: vi.fn(),
      visibleEditorValue: '',
      visibleSelectedCell: { sheetName: 'Sheet2', address: 'C3' },
      visibleSelection: { sheetName: 'Sheet2', address: 'C3' },
      workbookReady: true,
      workerHandle: {
        viewportStore: {},
      },
      writesAllowed: true,
      zeroConfigured: true,
    })

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: 'guest:test',
            defaultDocumentId: 'doc-1',
            persistState: true,
            zeroCacheUrl: 'http://127.0.0.1:4848',
          }}
          connectionState={{ name: 'connected' }}
        />,
      )
    })

    const onDeleteRows = latestWorkbookViewProps.current?.['onDeleteRows']
    const onDeleteColumns = latestWorkbookViewProps.current?.['onDeleteColumns']
    const onColumnWidthChange = latestWorkbookViewProps.current?.['onColumnWidthChange']
    const onRowHeightChange = latestWorkbookViewProps.current?.['onRowHeightChange']
    expect(typeof onDeleteRows).toBe('function')
    expect(typeof onDeleteColumns).toBe('function')
    expect(typeof onColumnWidthChange).toBe('function')
    expect(typeof onRowHeightChange).toBe('function')
    if (
      typeof onDeleteRows !== 'function' ||
      typeof onDeleteColumns !== 'function' ||
      typeof onColumnWidthChange !== 'function' ||
      typeof onRowHeightChange !== 'function'
    ) {
      throw new Error('WorkbookView did not receive mutation handlers')
    }

    const returnedDeleteRowsTask = onDeleteRows(1, 2)
    const returnedDeleteColumnsTask = onDeleteColumns(3, 1)
    onColumnWidthChange(0, 152)
    onRowHeightChange(0, 48)

    expect(latestWorkbookViewProps.current?.['sheetName']).toBe('Sheet2')
    expect(latestWorkbookViewProps.current?.['selectedAddr']).toBe('C3')
    expect(latestWorkbookViewProps.current?.['selectedCellSnapshot']).toEqual({ sheetName: 'Sheet2', address: 'C3' })
    expect(invokeDeleteRowsMutation).toHaveBeenCalledWith('Sheet2', 1, 2)
    expect(invokeDeleteColumnsMutation).toHaveBeenCalledWith('Sheet2', 3, 1)
    expect(invokeColumnWidthMutation).toHaveBeenCalledWith('Sheet2', 0, 152, {
      deferLocalApplication: true,
      deferPersistence: true,
    })
    expect(invokeRowHeightMutation).toHaveBeenCalledWith('Sheet2', 0, 48, {
      deferLocalApplication: true,
      deferPersistence: true,
    })
    expect(returnedDeleteRowsTask).toBe(deleteRowsTask)
    expect(returnedDeleteColumnsTask).toBe(deleteColumnsTask)

    await act(async () => {
      await Promise.all([returnedDeleteRowsTask, returnedDeleteColumnsTask])
      root.unmount()
    })
  })

  it('switches sheets from Google Sheets-style alt arrow shortcuts without leaving the current address', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectAddress = vi.fn()
    useWorkerWorkbookAppState.mockReturnValue(
      createReadyWorkbookAppState({
        selectAddress,
        sheetIdsByName: { Sheet1: 1, Sheet2: 2, Sheet3: 3 },
        sheetNames: ['Sheet1', 'Sheet2', 'Sheet3'],
        sheetOrdinalsByName: { Sheet1: 0, Sheet2: 1, Sheet3: 2 },
        selection: { sheetName: 'Sheet2', address: 'C22' },
        selectionSnapshot: createCellSelectionSnapshot('Sheet2', 'C22'),
        selectedCell: { sheetName: 'Sheet2', address: 'C22' },
        visibleSelectedCell: { sheetName: 'Sheet2', address: 'C22' },
        visibleSelection: { sheetName: 'Sheet2', address: 'C22' },
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          <WorkerWorkbookApp
            config={{
              currentUserId: 'guest:test',
              defaultDocumentId: 'doc-1',
              persistState: true,
              zeroCacheUrl: 'http://127.0.0.1:4848',
            }}
            connectionState={{ name: 'connected' }}
          />,
        )
      })

      await act(async () => {
        window.dispatchEvent(new KeyboardEvent('keydown', { altKey: true, bubbles: true, cancelable: true, key: 'ArrowDown' }))
      })
      expect(selectAddress).toHaveBeenLastCalledWith('Sheet3', 'C22')

      await act(async () => {
        window.dispatchEvent(new KeyboardEvent('keydown', { altKey: true, bubbles: true, cancelable: true, key: 'ArrowUp' }))
      })
      expect(selectAddress).toHaveBeenLastCalledWith('Sheet1', 'C22')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })
})
