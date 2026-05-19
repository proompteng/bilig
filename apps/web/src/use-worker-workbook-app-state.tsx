import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { useActorRef, useSelector } from '@xstate/react'
import { isWorkbookAgentCommandBundle, isWorkbookAgentPreviewSummary, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, WorkbookMergeRangeSnapshot } from '@bilig/protocol'
import { createWorkerRuntimeMachine, getWorkerRuntimeController, getWorkerRuntimeHandle } from './runtime-machine.js'
import { createRuntimeFetch, type resolveRuntimeConfig } from './runtime-config.js'
import type { ZeroClient } from './runtime-session.js'
import { loadPersistedSelection } from './selection-persistence.js'
import { type ZeroConnectionState, canAttemptRemoteSync, emptyCellSnapshot } from './worker-workbook-app-model.js'
import { useWorkbookSync } from './use-workbook-sync.js'
import { useWorkbookToolbar } from './use-workbook-toolbar.js'
import { useZeroHealthReady } from './use-zero-health-ready.js'
import { useWorkbookAppPanels } from './use-workbook-app-panels.js'
import { useWorkbookChangesPane } from './use-workbook-changes-pane.js'
import { useWorkbookSheetActions } from './use-workbook-sheet-actions.js'
import { createWorkbookPerfSession } from './perf/workbook-perf.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import { registerRuntimeDisposalHandlers } from './runtime-disposal-handlers.js'
import { isInstallBenchmarkCorpusResult } from './benchmark-corpus-result.js'
import { loadOrCreateWorkbookPresenceClientId } from './workbook-presence-client.js'
import { loadOrCreateWorkbookReplicaId } from './workbook-replica-client.js'
import { useWorkerWorkbookAgentContext } from './use-worker-workbook-agent-context.js'
import { useWorkerWorkbookGridState } from './use-worker-workbook-grid-state.js'
import { useWorkerWorkbookInteractionState } from './use-worker-workbook-interaction-state.js'

const workerRuntimeMachine = createWorkerRuntimeMachine()

function formatColumnName(columnIndex: number): string {
  let columnName = ''
  let value = columnIndex
  do {
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26) - 1
  } while (value >= 0)
  return columnName
}

function cellRangesIntersect(left: CellRangeRef, right: WorkbookMergeRangeSnapshot): boolean {
  if (left.sheetName !== right.sheetName) {
    return false
  }
  const leftStart = parseCellAddress(left.startAddress, left.sheetName)
  const leftEnd = parseCellAddress(left.endAddress, left.sheetName)
  const rightStart = parseCellAddress(right.startAddress, right.sheetName)
  const rightEnd = parseCellAddress(right.endAddress, right.sheetName)
  const leftMinRow = Math.min(leftStart.row, leftEnd.row)
  const leftMaxRow = Math.max(leftStart.row, leftEnd.row)
  const leftMinCol = Math.min(leftStart.col, leftEnd.col)
  const leftMaxCol = Math.max(leftStart.col, leftEnd.col)
  const rightMinRow = Math.min(rightStart.row, rightEnd.row)
  const rightMaxRow = Math.max(rightStart.row, rightEnd.row)
  const rightMinCol = Math.min(rightStart.col, rightEnd.col)
  const rightMaxCol = Math.max(rightStart.col, rightEnd.col)
  return leftMinRow <= rightMaxRow && leftMaxRow >= rightMinRow && leftMinCol <= rightMaxCol && leftMaxCol >= rightMinCol
}

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
  const fetchImpl = useMemo(() => createRuntimeFetch(runtimeConfig.serverUrl), [runtimeConfig.serverUrl])
  const authoritativeSyncEnabled = zeroConfigured || Boolean(runtimeConfig.serverUrl)
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
  const replicaId = useMemo(
    () => loadOrCreateWorkbookReplicaId(documentId, runtimeConfig.currentUserId),
    [documentId, runtimeConfig.currentUserId],
  )
  const presenceClientId = useMemo(() => loadOrCreateWorkbookPresenceClientId(), [])
  const selectionPersistenceScope = useMemo(
    () => ({
      documentId,
      userId: runtimeConfig.currentUserId,
    }),
    [documentId, runtimeConfig.currentUserId],
  )
  const initialSelection = useMemo(() => loadPersistedSelection(selectionPersistenceScope), [selectionPersistenceScope])
  const perfSession = useMemo(() => createWorkbookPerfSession({ documentId }), [documentId])

  useEffect(() => {
    perfSession.markShellMounted()
  }, [perfSession])

  const runtimeActorRef = useActorRef(workerRuntimeMachine, {
    input: {
      documentId,
      replicaId,
      persistState: runtimeConfig.persistState,
      authoritativeSyncEnabled,
      authoritativeEventSyncEnabled: zeroConfigured,
      perfSession,
      connectionStateName: connectionState.name,
      fetchImpl,
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
  const clearRuntimeError = useCallback(() => {
    runtimeActorRef.send({ type: 'error.clear' })
  }, [runtimeActorRef])
  const {
    hasLocalMutationInFlight,
    invokeMutation: invokeWorkbookMutation,
    invokeColumnVisibilityMutation,
    invokeColumnWidthMutation,
    invokeRowHeightMutation,
    invokeRowVisibilityMutation,
    redoLocalChange,
    retryPendingMutation,
    undoLocalChange,
  } = useWorkbookSync({
    documentId,
    connectionStateName: connectionState.name,
    connectionStateRef,
    runtimeController,
    workerHandleRef,
    zeroRef,
    reportRuntimeError,
  })
  const [localMutationEpoch, setLocalMutationEpoch] = useState(0)
  const localMutationEpochRef = useRef(0)
  const invokeMutation: typeof invokeWorkbookMutation = useCallback(
    async (method, ...args) => {
      localMutationEpochRef.current += 1
      setLocalMutationEpoch((epoch) => epoch + 1)
      await invokeWorkbookMutation(method, ...args)
    },
    [invokeWorkbookMutation],
  )
  const {
    columnWidths,
    rowHeights,
    hiddenColumns,
    hiddenRows,
    freezeRows,
    freezeCols,
    mergeRanges,
    selectedCell,
    invokeInsertRowsMutation: invokeInsertRowsMutationBase,
    invokeDeleteRowsMutation: invokeDeleteRowsMutationBase,
    invokeInsertColumnsMutation: invokeInsertColumnsMutationBase,
    invokeDeleteColumnsMutation: invokeDeleteColumnsMutationBase,
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
    editorTargetSelection,
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
    supersedeOptimisticCellSeedsForSheet,
    toggleBooleanCell,
    visibleSelectedCell,
    visibleEditorValue,
    visibleResolvedValue,
    visibleSelection,
  } = useWorkerWorkbookInteractionState({
    documentId,
    currentUserId: runtimeConfig.currentUserId,
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
  const invokeSheetStructuralMutation = useCallback(
    (taskFactory: () => Promise<void>, sheetName: string): Promise<void> => {
      const rollbackOptimisticSeeds = supersedeOptimisticCellSeedsForSheet(sheetName)
      const task = taskFactory()
      void (async () => {
        try {
          await task
        } catch {
          rollbackOptimisticSeeds?.()
        }
      })()
      return task
    },
    [supersedeOptimisticCellSeedsForSheet],
  )
  const invokeInsertRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> =>
      invokeSheetStructuralMutation(() => invokeInsertRowsMutationBase(sheetName, startRow, count), sheetName),
    [invokeInsertRowsMutationBase, invokeSheetStructuralMutation],
  )
  const invokeDeleteRowsMutation = useCallback(
    (sheetName: string, startRow: number, count: number): Promise<void> =>
      invokeSheetStructuralMutation(() => invokeDeleteRowsMutationBase(sheetName, startRow, count), sheetName),
    [invokeDeleteRowsMutationBase, invokeSheetStructuralMutation],
  )
  const invokeInsertColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> =>
      invokeSheetStructuralMutation(() => invokeInsertColumnsMutationBase(sheetName, startCol, count), sheetName),
    [invokeInsertColumnsMutationBase, invokeSheetStructuralMutation],
  )
  const invokeDeleteColumnsMutation = useCallback(
    (sheetName: string, startCol: number, count: number): Promise<void> =>
      invokeSheetStructuralMutation(() => invokeDeleteColumnsMutationBase(sheetName, startCol, count), sheetName),
    [invokeDeleteColumnsMutationBase, invokeSheetStructuralMutation],
  )
  const resolvedValue = visibleResolvedValue
  const { agentContextVersion, getAgentContext, handleVisibleViewportChange, resetVisibleViewportForSheet } = useWorkerWorkbookAgentContext(
    {
      selection,
      selectionRangeRef,
      selectionSnapshotRef,
      selectionRef,
      workerHandleRef,
      runtimeControllerRef,
    },
  )
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

  const {
    canRedo: remoteCanRedo,
    canUndo: remoteCanUndo,
    changeCount,
    changesPanel,
    redoLatestChange: redoRemoteLatestChange,
    undoLatestChange: undoRemoteLatestChange,
  } = useWorkbookChangesPane({
    documentId,
    currentUserId: runtimeConfig.currentUserId,
    sheetNames,
    zero: zeroSource,
    enabled: runtimeReady && zeroConfigured,
    pendingMutationSummary: runtimeState?.pendingMutationSummary,
    localMutationEpoch,
    localMutationEpochRef,
    onHistoryMutationApplied: async () => {
      await runtimeController?.invoke('refreshAuthoritativeEvents')
    },
    onJump: (sheetName, address) => {
      selectAddress(sheetName, address)
    },
  })
  const localHistoryState = runtimeState?.localHistoryState ?? {
    canUndo: false,
    canRedo: false,
  }
  const undoLocalLatestChange = useCallback(() => {
    void (async () => {
      try {
        await undoLocalChange()
      } catch (error) {
        reportRuntimeError(error)
      }
    })()
  }, [reportRuntimeError, undoLocalChange])
  const redoLocalLatestChange = useCallback(() => {
    void (async () => {
      try {
        await redoLocalChange()
      } catch (error) {
        reportRuntimeError(error)
      }
    })()
  }, [redoLocalChange, reportRuntimeError])
  const canUndo = zeroConfigured ? remoteCanUndo : !hasLocalMutationInFlight && localHistoryState.canUndo
  const canRedo = zeroConfigured ? remoteCanRedo : !hasLocalMutationInFlight && localHistoryState.canRedo
  const undoLatestChange = zeroConfigured ? undoRemoteLatestChange : undoLocalLatestChange
  const redoLatestChange = zeroConfigured ? redoRemoteLatestChange : redoLocalLatestChange

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

  const selectedStyle = workerHandle?.viewportStore.getCellStyle(visibleSelectedCell.styleId)
  const selectedPosition = useMemo(
    () => parseCellAddress(visibleSelection.address, visibleSelection.sheetName),
    [visibleSelection.address, visibleSelection.sheetName],
  )
  const selectedRange = selectionRangeRef.current
  const canUnmergeSelection = mergeRanges.some((mergeRange) => cellRangesIntersect(selectedRange, mergeRange))
  const failedPendingMutation = runtimeState?.pendingMutationSummary?.firstFailed ?? null
  const runtimeSyncState = runtimeState?.syncState ?? (runtimeReady ? 'syncing' : 'local-only')

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
    zeroHealthReady,
    changeCount,
    changesPanel,
    selectAddress,
    getAgentContext,
    agentContextVersion,
    applyAgentContext,
    previewAgentCommandBundle,
    syncAgentAuthoritativeRevision,
  })

  const { ribbon, statusModeLabel } = useWorkbookToolbar({
    connectionStateName: connectionState.name,
    runtimeReady,
    pendingMutationSummary: runtimeState?.pendingMutationSummary,
    failedPendingMutation,
    remoteSyncAvailable,
    zeroConfigured,
    zeroHealthReady,
    hasLocalMutationInFlight,
    canRedo,
    canUndo,
    canHideCurrentColumn: hiddenColumns[selectedPosition.col] !== true,
    canHideCurrentRow: hiddenRows[selectedPosition.row] !== true,
    canUnhideCurrentColumn: hiddenColumns[selectedPosition.col] === true,
    canUnhideCurrentRow: hiddenRows[selectedPosition.row] === true,
    canUnmergeSelection,
    invokeMutation,
    onHideCurrentColumn: () => {
      void (async () => {
        try {
          await invokeColumnVisibilityMutation(visibleSelection.sheetName, selectedPosition.col, true)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    onHideCurrentRow: () => {
      void (async () => {
        try {
          await invokeRowVisibilityMutation(visibleSelection.sheetName, selectedPosition.row, true)
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
          await invokeColumnVisibilityMutation(visibleSelection.sheetName, selectedPosition.col, false)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    onUnhideCurrentRow: () => {
      void (async () => {
        try {
          await invokeRowVisibilityMutation(visibleSelection.sheetName, selectedPosition.row, false)
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    selectionRangeRef,
    selectedCell: visibleSelectedCell,
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
      const result = await runtimeController.invoke('installBenchmarkCorpus', corpusId)
      if (!isInstallBenchmarkCorpusResult(result)) {
        throw new Error('Benchmark corpus install returned an unexpected payload')
      }
      const selectionAddress = `${formatColumnName(result.primaryViewport.colStart)}${String(result.primaryViewport.rowStart + 1)}`
      selectAddress(result.primaryViewport.sheetName, selectionAddress)
      getWorkbookScrollPerfCollector()?.setFixture({
        id: result.id,
        materializedCellCount: result.materializedCellCount,
        sheetName: result.primaryViewport.sheetName,
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
    editorTargetSelection,
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
    runtimeSyncState,
    retryFailedPendingMutation,
    getCellEditorSeed,
    selectAddress,
    selectSelectionSnapshot,
    selectedCell,
    selection,
    selectionSnapshot,
    visibleSelectedCell,
    visibleSelection,
    sidePanelId,
    acknowledgeExternalSelectionSync,
    handleSelectionChange,
    setSidePanelWidth,
    sheetIdsByName,
    sheetOrdinalsByName,
    sheetNames,
    sidePanel,
    sidePanelWidth,
    statusModeLabel,
    toggleBooleanCell,
    visibleEditorValue,
    workbookReady,
    workerHandle,
    writesAllowed,
    zeroConfigured,
  }
}
