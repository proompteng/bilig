// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkerWorkbookApp } from '../WorkerWorkbookApp.js'

interface ProbeGridEngine {
  getCell(sheetName: string, address: string): CellSnapshot
}

const { latestWorkbookViewProps, useWorkerWorkbookAppState } = vi.hoisted(() => ({
  latestWorkbookViewProps: { current: null as Record<string, unknown> | null },
  useWorkerWorkbookAppState: vi.fn(),
}))

vi.mock('@bilig/grid', () => ({
  WorkbookView: (props: Record<string, unknown>) => {
    latestWorkbookViewProps.current = props
    const engine = props['engine']
    if (!isProbeGridEngine(engine)) {
      throw new Error('Expected WorkbookView engine probe')
    }
    const sheetName = String(props['sheetName'])
    const selectedAddr = String(props['selectedAddr'])
    const snapshot = engine.getCell(sheetName, selectedAddr)
    const visibleText = snapshot.value.tag === ValueTag.String ? snapshot.value.value : ''
    return (
      <div data-testid="sheet-grid">
        <span data-testid="visible-cell-text">{visibleText}</span>
      </div>
    )
  },
}))

function isProbeGridEngine(value: unknown): value is ProbeGridEngine {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'getCell') === 'function'
}

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

describe('workbook visible-state regressions', () => {
  afterEach(() => {
    vi.clearAllMocks()
    latestWorkbookViewProps.current = null
    document.body.innerHTML = ''
  })

  it('renders authoritative Prepaid Template deep-link data through the visible grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const engine = createProbeGridEngine({
      'Prepaid Template!E46': createStringCellSnapshot('Prepaid Template', 'E46', 'Expense Recognized'),
    })
    useWorkerWorkbookAppState.mockReturnValue(
      createWorkbookAppState({
        selectedCell: { sheetName: 'Prepaid Template', address: 'E46' },
        selection: { sheetName: 'Prepaid Template', address: 'E46' },
        selectionSnapshot: {
          sheetName: 'Prepaid Template',
          address: 'E46',
          kind: 'cell',
          range: {
            startAddress: 'E46',
            endAddress: 'E46',
          },
        },
        visibleSelectedCell: { sheetName: 'Prepaid Template', address: 'E46' },
        visibleSelection: { sheetName: 'Prepaid Template', address: 'E46' },
        workerHandle: { viewportStore: engine },
      }),
    )

    const { host, root } = await renderWorkbookApp()

    expect(host.querySelector("[data-testid='missing-sheet-state']")).toBeNull()
    expect(latestWorkbookViewProps.current?.['sheetName']).toBe('Prepaid Template')
    expect(latestWorkbookViewProps.current?.['selectedAddr']).toBe('E46')
    expect(host.querySelector("[data-testid='visible-cell-text']")?.textContent).toBe('Expense Recognized')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders a typed missing-sheet state instead of a successful empty viewport', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const selectAddress = vi.fn()
    useWorkerWorkbookAppState.mockReturnValue(
      createWorkbookAppState({
        selectAddress,
        selectedCell: { sheetName: 'Missing Sheet', address: 'E46' },
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
        visibleSelectedCell: { sheetName: 'Missing Sheet', address: 'E46' },
        visibleSelection: { sheetName: 'Missing Sheet', address: 'E46' },
      }),
    )

    const { host, root } = await renderWorkbookApp()

    expect(host.querySelector("[data-testid='missing-sheet-state']")?.textContent).toContain('Missing Sheet')
    expect(latestWorkbookViewProps.current).toBeNull()

    const recoveryButton = [...host.querySelectorAll('button')].find((button) => button.textContent === 'Open Prepaid Template')
    expect(recoveryButton).toBeInstanceOf(HTMLButtonElement)
    if (!(recoveryButton instanceof HTMLButtonElement)) {
      throw new Error('Expected missing-sheet recovery button')
    }

    await act(async () => {
      recoveryButton.click()
    })

    expect(selectAddress).toHaveBeenCalledWith('Prepaid Template', 'A1')

    await act(async () => {
      root.unmount()
    })
  })

  it('holds a resolving state for URL-requested sheets while workbook sync is still loading', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    useWorkerWorkbookAppState.mockReturnValue(
      createWorkbookAppState({
        selectedCell: { sheetName: 'Prepaid Template', address: 'C12' },
        selection: { sheetName: 'Prepaid Template', address: 'C12' },
        selectionSnapshot: {
          sheetName: 'Prepaid Template',
          address: 'C12',
          kind: 'cell',
          range: {
            startAddress: 'C12',
            endAddress: 'C12',
          },
        },
        sheetIdsByName: { Sheet1: 1 },
        sheetNames: ['Sheet1'],
        sheetOrdinalsByName: { Sheet1: 0 },
        runtimeSyncState: 'syncing',
        visibleSelectedCell: { sheetName: 'Prepaid Template', address: 'C12' },
        visibleSelection: { sheetName: 'Prepaid Template', address: 'C12' },
      }),
    )

    const { host, root } = await renderWorkbookApp()

    expect(host.querySelector("[data-testid='missing-sheet-state']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-resolving-state']")?.textContent).toContain('Prepaid Template')
    expect(latestWorkbookViewProps.current).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })
})

function createWorkbookAppState(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  const engine = createProbeGridEngine({})
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
    installBenchmarkCorpus: vi.fn(),
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
    selection: { sheetName: 'Prepaid Template', address: 'E46' },
    selectionSnapshot: {
      sheetName: 'Prepaid Template',
      address: 'E46',
      kind: 'cell',
      range: {
        startAddress: 'E46',
        endAddress: 'E46',
      },
    },
    selectAddress: vi.fn(),
    selectSelectionSnapshot: vi.fn(),
    selectedCell: { sheetName: 'Prepaid Template', address: 'E46' },
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
    visibleSelectedCell: { sheetName: 'Prepaid Template', address: 'E46' },
    visibleSelection: { sheetName: 'Prepaid Template', address: 'E46' },
    workbookReady: true,
    workerHandle: {
      viewportStore: engine,
    },
    writesAllowed: true,
    zeroConfigured: true,
    ...overrides,
  }
}

function createProbeGridEngine(cells: Record<string, CellSnapshot>): ProbeGridEngine {
  return {
    getCell(sheetName, address) {
      return cells[`${sheetName}!${address}`] ?? createEmptyCellSnapshot(sheetName, address)
    },
  }
}

function createStringCellSnapshot(sheetName: string, address: string, value: string): CellSnapshot {
  return {
    sheetName,
    address,
    flags: 0,
    input: value,
    value: {
      tag: ValueTag.String,
      value,
      stringId: 1,
    },
    version: 1,
  }
}

function createEmptyCellSnapshot(sheetName: string, address: string): CellSnapshot {
  return {
    sheetName,
    address,
    flags: 0,
    value: {
      tag: ValueTag.Empty,
    },
    version: 1,
  }
}

async function renderWorkbookApp() {
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

  return { host, root }
}
