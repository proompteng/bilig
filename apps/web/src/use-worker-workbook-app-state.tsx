import { useCallback, useEffect, useMemo, useRef, type ReactNode } from 'react'
import { useActorRef, useSelector } from '@xstate/react'
import { isWorkbookAgentCommandBundle, isWorkbookAgentPreviewSummary, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { parseCellAddress } from '@bilig/formula'
import { createWorkerRuntimeMachine, getWorkerRuntimeController, getWorkerRuntimeHandle } from './runtime-machine.js'
import type { resolveRuntimeConfig } from './runtime-config.js'
import type { ZeroClient } from './runtime-session.js'
import { loadPersistedSelection } from './selection-persistence.js'
import { type ZeroConnectionState, canAttemptRemoteSync, emptyCellSnapshot, toResolvedValue } from './worker-workbook-app-model.js'
import { useWorkbookSync } from './use-workbook-sync.js'
import { useWorkbookToolbar } from './use-workbook-toolbar.js'
import { useZeroHealthReady } from './use-zero-health-ready.js'
import { useWorkbookAppPanels } from './use-workbook-app-panels.js'
import { useWorkbookChangesPane } from './use-workbook-changes-pane.js'
import { useWorkbookSheetActions } from './use-workbook-sheet-actions.js'
import { createWorkbookPerfSession } from './perf/workbook-perf.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import { registerRuntimeDisposalHandlers } from './runtime-disposal-handlers.js'
import { useWorkbookLocalPersistenceHandoff } from './use-workbook-local-persistence-handoff.js'
import { loadOrCreateWorkbookPresenceClientId } from './workbook-presence-client.js'
import { useWorkerWorkbookAgentContext } from './use-worker-workbook-agent-context.js'
import { useWorkerWorkbookGridState } from './use-worker-workbook-grid-state.js'
import { useWorkerWorkbookInteractionState } from './use-worker-workbook-interaction-state.js'

const workerRuntimeMachine = createWorkerRuntimeMachine()

interface LocalOnlyZeroSource {
  materialize(query: unknown): {
    data: unknown
    addListener(listener: (value: unknown) => void): () => void
    destroy(): void
  }
  mutate(mutation: unknown): unknown
}

export function useWorkerWorkbookAppState(input: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>
  connectionState: ZeroConnectionState
  zero?: ZeroClient
  toolbarControls?: ReactNode
}) {
  const { runtimeConfig, connectionState, zero, toolbarControls } = input
  const documentId = runtimeConfig.documentId
  const zeroConfigured = Boolean(zero)
  const zeroSource = useMemo<LocalOnlyZeroSource>(
    () =>
      zero ??
      ({
        materialize(_query: unknown) {
          return {
            data: [],
            addListener(_listener: (value: unknown) => void) {
              return () => undefined
            },
            destroy() {},
          }
        },
        mutate(_mutation: unknown) {
          return {}
        },
      } satisfies LocalOnlyZeroSource),
    [zero],
  )
  const replicaId = useMemo(() => `browser:${Math.random().toString(36).slice(2)}`, [])
  const presenceClientId = useMemo(() => loadOrCreateWorkbookPresenceClientId(), [])
  const initialSelection = useMemo(() => loadPersistedSelection(documentId), [documentId])
  const perfSession = useMemo(() => createWorkbookPerfSession({ documentId }), [documentId])

  useEffect(() => {
    perfSession.markShellMounted()
  }, [perfSession])

  const runtimeActorRef = useActorRef(workerRuntimeMachine, {
    input: {
      documentId,
      replicaId,
      persistState: runtimeConfig.persistState,
      perfSession,
      connectionStateName: connectionState.name,
      ...(zero ? { zero } : {}),
      initialSelection,
    },
  })
  const runtimeController = useSelector(runtimeActorRef, (snapshot) => getWorkerRuntimeController(snapshot.context))
  const workerHandle = useSelector(runtimeActorRef, (snapshot) => getWorkerRuntimeHandle(snapshot.context))
  const runtimeState = useSelector(runtimeActorRef, (snapshot) => snapshot.context.runtimeState)
  const selection = useSelector(runtimeActorRef, (snapshot) => snapshot.context.selection)
  const runtimeError = useSelector(runtimeActorRef, (snapshot) => snapshot.context.error)
  const runtimeReady = Boolean(workerHandle)
  const workbookReady = runtimeReady
  const emptySelectedCell = useMemo(
    () => emptyCellSnapshot(selection.sheetName, selection.address),
    [selection.address, selection.sheetName],
  )
  const workerHandleRef = useRef(workerHandle)
  const runtimeControllerRef = useRef(runtimeController)
  const zeroRef = useRef<LocalOnlyZeroSource>(zeroSource)
  const connectionStateRef = useRef(connectionState.name)

  workerHandleRef.current = workerHandle
  runtimeControllerRef.current = runtimeController
  zeroRef.current = zeroSource
  connectionStateRef.current = connectionState.name

  useEffect(() => {
    workerHandleRef.current = workerHandle
  }, [workerHandle])

  useEffect(() => {
    runtimeControllerRef.current = runtimeController
  }, [runtimeController])

  useEffect(() => {
    zeroRef.current = zeroSource
  }, [zeroSource])

  useEffect(() => {
    connectionStateRef.current = connectionState.name
  }, [connectionState.name])

  useEffect(() => {
    return registerRuntimeDisposalHandlers({
      getController: () => runtimeControllerRef.current,
    })
  }, [])

  useEffect(() => {
    runtimeActorRef.send({
      type: 'connection.changed',
      connectionStateName: connectionState.name,
    })
  }, [connectionState.name, runtimeActorRef])

  const writesAllowed = runtimeReady
  const remoteSyncAvailable = canAttemptRemoteSync(connectionState.name)
  const zeroHealthReady = useZeroHealthReady({
    connectionStateName: connectionState.name,
    runtimeReady,
  })

  const reportRuntimeError = useCallback(
    (error: unknown) => {
      runtimeActorRef.send({
        type: 'session.error',
        message: error instanceof Error ? error.message : String(error),
      })
    },
    [runtimeActorRef],
  )
  const retryRuntime = useCallback(
    (persistState: boolean) => {
      runtimeActorRef.send({ type: 'retry', persistState })
    },
    [runtimeActorRef],
  )
  const clearRuntimeError = useCallback(() => {
    runtimeActorRef.send({ type: 'error.clear' })
  }, [runtimeActorRef])
  const {
    invokeMutation,
    invokeColumnVisibilityMutation,
    invokeColumnWidthMutation,
    invokeRowHeightMutation,
    invokeRowVisibilityMutation,
    retryPendingMutation,
  } = useWorkbookSync({
    documentId,
    connectionStateName: connectionState.name,
    connectionStateRef,
    runtimeController,
    workerHandleRef,
    zeroRef,
    reportRuntimeError,
  })
  const {
    columnWidths,
    rowHeights,
    hiddenColumns,
    hiddenRows,
    freezeRows,
    freezeCols,
    selectedCell,
    invokeInsertRowsMutation,
    invokeDeleteRowsMutation,
    invokeInsertColumnsMutation,
    invokeDeleteColumnsMutation,
    invokeSetFreezePaneMutation,
  } = useWorkerWorkbookGridState({
    workerHandle,
    selection,
    emptySelectedCell,
    invokeMutation,
  })
  const sendSelectionChanged = useCallback(
    (nextSelection: typeof selection) => {
      runtimeActorRef.send({ type: 'selection.changed', selection: nextSelection })
    },
    [runtimeActorRef],
  )
  const {
    beginEditing,
    cancelEditor,
    clearSelectedCell,
    commitEditor,
    copySelectionRange,
    editorConflictBanner,
    editorSelectionBehavior,
    fillSelectionRange,
    getCellEditorSeed,
    acknowledgeExternalSelectionSync,
    handleEditorChange,
    handleSelectionChange,
    isEditing,
    isEditingCell,
    moveSelectionRange,
    pasteIntoSelection,
    selectAddress,
    selectionRangeRef,
    selectionRef,
    selectionSnapshot,
    selectionSnapshotRef,
    selectSelectionSnapshot,
    toggleBooleanCell,
    visibleEditorValue,
  } = useWorkerWorkbookInteractionState({
    documentId,
    selection,
    selectedCell,
    workerHandle,
    workerHandleRef,
    writesAllowed,
    invokeMutation,
    perfSession,
    reportRuntimeError,
    sendSelectionChanged,
  })
  const resolvedValue = toResolvedValue(selectedCell)
  const { getAgentContext, handleVisibleViewportChange, resetVisibleViewportForSheet } = useWorkerWorkbookAgentContext({
    selection,
    selectionRangeRef,
    selectionSnapshotRef,
    selectionRef,
    workerHandleRef,
    runtimeControllerRef,
  })
  const autofitColumn = useCallback(
    async (sheetName: string, columnIndex: number, fallbackWidth: number) => {
      const nextWidth =
        runtimeController && typeof runtimeController.invoke === 'function'
          ? await runtimeController.invoke('autofitColumn', sheetName, columnIndex)
          : fallbackWidth
      const normalizedWidth = typeof nextWidth === 'number' && Number.isFinite(nextWidth) ? nextWidth : fallbackWidth
      await invokeColumnWidthMutation(sheetName, columnIndex, normalizedWidth, {
        flush: true,
      })
    },
    [invokeColumnWidthMutation, runtimeController],
  )
  const sheetNames = useMemo(
    () => [...(runtimeState?.sheetNames ?? [selection.sheetName])],
    [runtimeState?.sheetNames, selection.sheetName],
  )
  const sheetIdsByName = useMemo(() => {
    const entries = runtimeState?.sheets?.map((sheet) => [sheet.name, sheet.id] as const) ?? []
    return Object.fromEntries(entries)
  }, [runtimeState?.sheets])
  const sheetOrdinalsByName = useMemo(() => {
    const entries = runtimeState?.sheets?.map((sheet) => [sheet.name, sheet.order] as const) ?? []
    return Object.fromEntries(entries)
  }, [runtimeState?.sheets])
  const definedNames = useMemo(() => [...(runtimeState?.definedNames ?? [])], [runtimeState])
  const previousViewportResetSheetNameRef = useRef<string | null>(null)
  useEffect(() => {
    if (previousViewportResetSheetNameRef.current === selection.sheetName) {
      return
    }
    previousViewportResetSheetNameRef.current = selection.sheetName
    resetVisibleViewportForSheet(selection)
  }, [resetVisibleViewportForSheet, selection])
  const applyAgentContext = useCallback(
    (context: ReturnType<typeof getAgentContext>) => {
      const range = context.selection.range ?? {
        startAddress: context.selection.address,
        endAddress: context.selection.address,
      }
      selectSelectionSnapshot({
        sheetName: context.selection.sheetName,
        address: context.selection.address,
        kind: range.startAddress === range.endAddress ? 'cell' : 'range',
        range,
      })
    },
    [selectSelectionSnapshot],
  )

  const { canRedo, canUndo, changeCount, changesPanel, redoLatestChange, undoLatestChange } = useWorkbookChangesPane({
    documentId,
    currentUserId: runtimeConfig.currentUserId,
    sheetNames,
    zero: zeroSource,
    enabled: runtimeReady && zeroConfigured,
    onJump: (sheetName, address) => {
      selectAddress(sheetName, address)
    },
  })

  const previewAgentCommandBundle = useCallback(
    async (bundle: WorkbookAgentCommandBundle) => {
      if (!runtimeController || !isWorkbookAgentCommandBundle(bundle)) {
        throw new Error('Workbook runtime is not ready for agent preview')
      }
      const value = await runtimeController.invoke('previewAgentCommandBundle', bundle)
      if (!isWorkbookAgentPreviewSummary(value)) {
        throw new Error('Worker returned an invalid workbook agent preview')
      }
      return value
    },
    [runtimeController],
  )
  const syncAgentAuthoritativeRevision = useCallback(
    async (revision: number) => {
      await runtimeController?.invoke('refreshAuthoritativeEvents', revision)
    },
    [runtimeController],
  )

  const selectedStyle = workerHandle?.viewportStore.getCellStyle(selectedCell.styleId)
  const selectedPosition = useMemo(() => parseCellAddress(selection.address, selection.sheetName), [selection.address, selection.sheetName])
  const failedPendingMutation = runtimeState?.pendingMutationSummary?.firstFailed ?? null
  const localPersistenceMode = runtimeState?.localPersistenceMode ?? 'ephemeral'
  const {
    approvePersistenceTransfer,
    dismissPersistenceTransferRequest,
    pendingTransferRequest,
    requestPersistenceTransfer,
    transferRequested,
  } = useWorkbookLocalPersistenceHandoff({
    documentId,
    localPersistenceMode,
    retryRuntime,
  })

  const {
    agentError,
    clearAgentError,
    agentPanel,
    previewRanges,
    sidePanelId,
    setSidePanelWidth,
    sidePanel,
    sidePanelWidth,
    toolbarTrailingContent,
  } = useWorkbookAppPanels({
    currentUserId: runtimeConfig.currentUserId,
    documentId,
    presenceClientId,
    replicaId,
    selection,
    sheetNames,
    zero: zeroSource,
    runtimeReady,
    zeroConfigured,
    workbookAgentEnabled: runtimeConfig.workbookAgentEnabled,
    remoteSyncAvailable,
    changeCount,
    changesPanel,
    selectAddress,
    getAgentContext,
    applyAgentContext,
    previewAgentCommandBundle,
    syncAgentAuthoritativeRevision,
  })

  const { ribbon, statusModeLabel } = useWorkbookToolbar({
    connectionStateName: connectionState.name,
    runtimeReady,
    localPersistenceMode,
    pendingMutationSummary: runtimeState?.pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    canRedo,
    canUndo,
    canHideCurrentColumn: hiddenColumns[selectedPosition.col] !== true,
    canHideCurrentRow: hiddenRows[selectedPosition.row] !== true,
    canUnhideCurrentColumn: hiddenColumns[selectedPosition.col] === true,
    canUnhideCurrentRow: hiddenRows[selectedPosition.row] === true,
    invokeMutation,
    onHideCurrentColumn: () => {
      void (async () => {
        try {
          await invokeColumnVisibilityMutation(selection.sheetName, selectedPosition.col, true)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    onHideCurrentRow: () => {
      void (async () => {
        try {
          await invokeRowVisibilityMutation(selection.sheetName, selectedPosition.row, true)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    onRedo: redoLatestChange,
    onUndo: undoLatestChange,
    onUnhideCurrentColumn: () => {
      void (async () => {
        try {
          await invokeColumnVisibilityMutation(selection.sheetName, selectedPosition.col, false)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    onUnhideCurrentRow: () => {
      void (async () => {
        try {
          await invokeRowVisibilityMutation(selection.sheetName, selectedPosition.row, false)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    selectionRangeRef,
    selectedCell,
    selectedStyle,
    trailingContent: (
      <>
        {toolbarControls}
        {toolbarTrailingContent}
      </>
    ),
    writesAllowed,
  })

  const { createSheet, deleteSheet, renameSheet } = useWorkbookSheetActions({
    sheetNames,
    selectionRef,
    invokeMutation,
    selectAddress,
    reportRuntimeError,
  })

  const installBenchmarkCorpus = useCallback(
    async (corpusId: string): Promise<void> => {
      if (!runtimeController) {
        throw new Error('Workbook runtime is not ready for benchmark install')
      }
      getWorkbookScrollPerfCollector()?.setBenchmarkState('loading')
      const benchmarks = await import('@bilig/benchmarks/workbook-corpus')
      if (!benchmarks.isWorkbookBenchmarkCorpusId(corpusId)) {
        throw new Error(`Unknown benchmark corpus ${corpusId}`)
      }
      const corpus = benchmarks.buildWorkbookBenchmarkCorpus(corpusId)
      await runtimeController.invoke('installAuthoritativeSnapshot', {
        snapshot: corpus.snapshot,
        authoritativeRevision: 0,
        mode: 'bootstrap',
      })
      await runtimeController.invoke('materializeProjectionEngine')
      const selectionAddress = `${String.fromCharCode(65 + corpus.primaryViewport.colStart)}${String(corpus.primaryViewport.rowStart + 1)}`
      selectAddress(corpus.primaryViewport.sheetName, selectionAddress)
      getWorkbookScrollPerfCollector()?.setFixture({
        id: corpus.id,
        materializedCellCount: corpus.materializedCellCount,
        sheetName: corpus.primaryViewport.sheetName,
      })
    },
    [runtimeController, selectAddress],
  )

  const retryFailedPendingMutation = useCallback(async (): Promise<void> => {
    if (!failedPendingMutation) {
      return
    }
    await retryPendingMutation(failedPendingMutation.id)
  }, [failedPendingMutation, retryPendingMutation])

  return {
    agentError,
    clearAgentError,
    clearRuntimeError,
    agentPanel,
    beginEditing,
    autofitColumn,
    cancelEditor,
    clearSelectedCell,
    columnWidths,
    hiddenColumns,
    hiddenRows,
    freezeCols,
    freezeRows,
    rowHeights,
    commitEditor,
    copySelectionRange,
    createSheet,
    changesPanel,
    deleteSheet,
    definedNames,
    editorConflictBanner,
    editorSelectionBehavior,
    failedPendingMutation,
    fillSelectionRange,
    handleEditorChange,
    handleVisibleViewportChange,
    invokeDeleteColumnsMutation,
    invokeDeleteRowsMutation,
    invokeInsertColumnsMutation,
    invokeInsertRowsMutation,
    invokeSetFreezePaneMutation,
    installBenchmarkCorpus,
    invokeColumnVisibilityMutation,
    invokeColumnWidthMutation,
    invokeRowHeightMutation,
    invokeRowVisibilityMutation,
    isEditing,
    isEditingCell,
    moveSelectionRange,
    pasteIntoSelection,
    previewRanges,
    remoteSyncAvailable,
    renameSheet,
    reportRuntimeError,
    resolvedValue,
    ribbon,
    runtimeError,
    runtimeReady,
    retryFailedPendingMutation,
    getCellEditorSeed,
    approvePersistenceTransfer,
    selectAddress,
    selectSelectionSnapshot,
    selectedCell,
    selection,
    selectionSnapshot,
    sidePanelId,
    dismissPersistenceTransferRequest,
    acknowledgeExternalSelectionSync,
    handleSelectionChange,
    setSidePanelWidth,
    sheetIdsByName,
    sheetOrdinalsByName,
    sheetNames,
    sidePanel,
    sidePanelWidth,
    statusModeLabel,
    localPersistenceMode,
    pendingTransferRequest,
    requestPersistenceTransfer,
    transferRequested,
    toggleBooleanCell,
    visibleEditorValue,
    workbookReady,
    workerHandle,
    writesAllowed,
    zeroConfigured,
  }
}
