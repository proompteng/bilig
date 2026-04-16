import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore, type ReactNode } from 'react'
import { useActorRef, useSelector } from '@xstate/react'
import { isWorkbookAgentCommandBundle, isWorkbookAgentPreviewSummary, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { EditMovement, EditSelectionBehavior } from '@bilig/grid'
import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, CellSnapshot, Viewport } from '@bilig/protocol'
import { createWorkerRuntimeMachine } from './runtime-machine.js'
import { resolveRuntimeConfig } from './runtime-config.js'
import type { ZeroClient } from './runtime-session.js'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import { loadPersistedSelection, persistSelection } from './selection-persistence.js'
import { ProjectedViewportStore } from './projected-viewport-store.js'
import { buildWorkbookAgentContext, singleCellAgentSelectionRange, type WorkbookAgentSelectionRange } from './workbook-agent-context.js'
import {
  type EditingMode,
  type ParsedEditorInput,
  type WorkbookEditorConflict,
  type ZeroConnectionState,
  canAttemptRemoteSync,
  clampSelectionMovement,
  emptyCellSnapshot,
  parseEditorInput,
  parsedEditorInputEquals,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
  toEditorValue,
  toResolvedValue,
} from './worker-workbook-app-model.js'
import {
  readViewportColumnWidths,
  readViewportHiddenColumns,
  readViewportHiddenRows,
  readViewportRowHeights,
} from './worker-workbook-view-state.js'
import { useWorkbookSync } from './use-workbook-sync.js'
import { useWorkbookToolbar } from './use-workbook-toolbar.js'
import { useZeroHealthReady } from './use-zero-health-ready.js'
import { useWorkbookAppPanels } from './use-workbook-app-panels.js'
import { useWorkbookChangesPane } from './use-workbook-changes-pane.js'
import { useWorkbookSheetActions } from './use-workbook-sheet-actions.js'
import { useWorkbookSelectionActions } from './use-workbook-selection-actions.js'
import { useWorkbookEditorConflict } from './use-workbook-editor-conflict.js'
import { createWorkbookPerfSession } from './perf/workbook-perf.js'
import { registerRuntimeDisposalHandlers } from './runtime-disposal-handlers.js'
import { useWorkbookLocalPersistenceHandoff } from './use-workbook-local-persistence-handoff.js'
import { loadOrCreateWorkbookPresenceClientId } from './workbook-presence-client.js'

const workerRuntimeMachine = createWorkerRuntimeMachine()

interface LocalOnlyZeroSource {
  materialize(query: unknown): {
    data: unknown
    addListener(listener: (value: unknown) => void): () => void
    destroy(): void
  }
  mutate(mutation: unknown): unknown
}

function selectionViewport(selection: WorkerRuntimeSelection): Viewport {
  const parsed = parseCellAddress(selection.address, selection.sheetName)
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  }
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
  const runtimeController = useSelector(runtimeActorRef, (snapshot) => snapshot.context.controller)
  const workerHandle = useSelector(runtimeActorRef, (snapshot) => snapshot.context.handle)
  const runtimeState = useSelector(runtimeActorRef, (snapshot) => snapshot.context.runtimeState)
  const selection = useSelector(runtimeActorRef, (snapshot) => snapshot.context.selection)
  const runtimeError = useSelector(runtimeActorRef, (snapshot) => snapshot.context.error)
  const runtimeReady = Boolean(workerHandle)
  const workbookReady = runtimeReady
  const emptySelectedCell = useMemo(
    () => emptyCellSnapshot(selection.sheetName, selection.address),
    [selection.address, selection.sheetName],
  )
  const [editorValue, setEditorValue] = useState('')
  const [editorSelectionBehavior, setEditorSelectionBehavior] = useState<EditSelectionBehavior>('select-all')
  const [editingMode, setEditingMode] = useState<EditingMode>('idle')
  const [editorConflict, setEditorConflict] = useState<WorkbookEditorConflict | null>(null)
  const selectionRef = useRef(selection)
  const workerHandleRef = useRef(workerHandle)
  const runtimeControllerRef = useRef(runtimeController)
  const editorValueRef = useRef(editorValue)
  const editingModeRef = useRef(editingMode)
  const editorTargetRef = useRef(selection)
  const editorBaseSnapshotRef = useRef<CellSnapshot>(emptySelectedCell)
  const zeroRef = useRef<LocalOnlyZeroSource>(zeroSource)
  const connectionStateRef = useRef(connectionState.name)
  const visibleViewportRef = useRef<Viewport>(selectionViewport(selection))
  const agentSelectionRangeRef = useRef<WorkbookAgentSelectionRange>(singleCellAgentSelectionRange(selection))
  const selectionRangeRef = useRef<CellRangeRef>({
    sheetName: selection.sheetName,
    startAddress: selection.address,
    endAddress: selection.address,
  })

  useEffect(() => {
    const previousSelection = selectionRef.current
    selectionRef.current = selection
    if (previousSelection.sheetName !== selection.sheetName || previousSelection.address !== selection.address) {
      selectionRangeRef.current = {
        sheetName: selection.sheetName,
        startAddress: selection.address,
        endAddress: selection.address,
      }
      agentSelectionRangeRef.current = singleCellAgentSelectionRange(selection)
    }
  }, [selection])

  useEffect(() => {
    persistSelection(documentId, selection)
  }, [documentId, selection])

  useEffect(() => {
    workerHandleRef.current = workerHandle
  }, [workerHandle])

  useEffect(() => {
    runtimeControllerRef.current = runtimeController
  }, [runtimeController])

  useEffect(() => {
    editorValueRef.current = editorValue
  }, [editorValue])

  useEffect(() => {
    editingModeRef.current = editingMode
  }, [editingMode])

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

  const columnWidths = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
  )

  const rowHeights = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
  )

  const hiddenColumns = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
  )

  const hiddenRows = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
  )

  const freezeRows = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
  )

  const freezeCols = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
  )

  const selectedCell = useSyncExternalStore(
    useCallback((listener: () => void) => workerHandle?.viewportStore.subscribe(listener) ?? (() => {}), [workerHandle]),
    () => workerHandle?.viewportStore.peekCell(selection.sheetName, selection.address) ?? emptySelectedCell,
    () => emptySelectedCell,
  )

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

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current
      if (!active) {
        return selectedCell
      }
      return active.viewportStore.getCell(nextSelection.sheetName, nextSelection.address)
    },
    [selectedCell],
  )

  const cloneLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => structuredClone(getLiveSelectedCell(nextSelection)),
    [getLiveSelectedCell],
  )

  const resetEditorConflictTracking = useCallback(
    (nextSelection = selectionRef.current) => {
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextSelection)
      setEditorConflict(null)
    },
    [cloneLiveSelectedCell],
  )

  const completeEditNavigation = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      if (!movement) {
        selectionRef.current = targetSelection
        editorTargetRef.current = targetSelection
        return targetSelection
      }
      const nextAddress = clampSelectionMovement(targetSelection.address, targetSelection.sheetName, movement)
      const nextSelection = { sheetName: targetSelection.sheetName, address: nextAddress }
      selectionRef.current = nextSelection
      editorTargetRef.current = nextSelection
      runtimeActorRef.send({ type: 'selection.changed', selection: nextSelection })
      return nextSelection
    },
    [runtimeActorRef],
  )

  const finishEditingWithAuthoritative = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      const nextSelection = completeEditNavigation(targetSelection, movement)
      const nextEditorValue = toEditorValue(cloneLiveSelectedCell(nextSelection))
      editorValueRef.current = nextEditorValue
      setEditorValue(nextEditorValue)
      setEditorSelectionBehavior('select-all')
      editingModeRef.current = 'idle'
      setEditingMode('idle')
      resetEditorConflictTracking(nextSelection)
    },
    [cloneLiveSelectedCell, completeEditNavigation, resetEditorConflictTracking],
  )

  const beginEditing = useCallback(
    (seed?: string, selectionBehavior: EditSelectionBehavior = 'select-all', mode: Exclude<EditingMode, 'idle'> = 'cell') => {
      if (!writesAllowed) {
        return
      }
      const nextEditorValue = seed ?? toEditorValue(getLiveSelectedCell())
      const nextTarget = selectionRef.current
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextTarget)
      setEditorConflict(null)
      editorValueRef.current = nextEditorValue
      setEditorValue(nextEditorValue)
      setEditorSelectionBehavior(selectionBehavior)
      editorTargetRef.current = nextTarget
      editingModeRef.current = mode
      setEditingMode(mode)
    },
    [cloneLiveSelectedCell, getLiveSelectedCell, writesAllowed],
  )

  const applyParsedInput = useCallback(
    async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      if (parsed.kind === 'formula') {
        await invokeMutation('setCellFormula', sheetName, address, parsed.formula)
        perfSession.markFirstLocalEditApplied?.()
        return
      }
      if (parsed.kind === 'clear') {
        await invokeMutation('clearCell', sheetName, address)
        perfSession.markFirstLocalEditApplied?.()
        return
      }
      await invokeMutation('setCellValue', sheetName, address, parsed.value)
      perfSession.markFirstLocalEditApplied?.()
    },
    [invokeMutation, perfSession],
  )

  const commitEditor = useCallback(
    (movement?: EditMovement) => {
      if (!writesAllowed) {
        return
      }
      const targetSelection = editingModeRef.current === 'idle' ? selectionRef.current : editorTargetRef.current
      const nextValue = editingModeRef.current === 'idle' ? toEditorValue(getLiveSelectedCell(targetSelection)) : editorValueRef.current
      const parsed = parseEditorInput(nextValue)
      const baseSnapshot = editorBaseSnapshotRef.current
      const authoritativeSnapshot = cloneLiveSelectedCell(targetSelection)
      const draftMatchesAuthoritative = parsedEditorInputMatchesSnapshot(parsed, authoritativeSnapshot)
      const draftMatchesBase = parsedEditorInputEquals(parsed, parsedEditorInputFromSnapshot(baseSnapshot))

      if (!sameCellContent(baseSnapshot, authoritativeSnapshot) && !draftMatchesAuthoritative) {
        if (draftMatchesBase) {
          finishEditingWithAuthoritative(targetSelection, movement)
          return
        }
        setEditorConflict({
          sheetName: targetSelection.sheetName,
          address: targetSelection.address,
          phase: 'compare',
          baseSnapshot,
          authoritativeSnapshot,
        })
        return
      }

      if (draftMatchesAuthoritative) {
        finishEditingWithAuthoritative(targetSelection, movement)
        return
      }

      const nextSelection = completeEditNavigation(targetSelection, movement)
      setEditorSelectionBehavior('select-all')
      editingModeRef.current = 'idle'
      setEditingMode('idle')
      resetEditorConflictTracking(nextSelection)
      void applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed).catch(reportRuntimeError)
    },
    [
      applyParsedInput,
      cloneLiveSelectedCell,
      completeEditNavigation,
      finishEditingWithAuthoritative,
      getLiveSelectedCell,
      reportRuntimeError,
      resetEditorConflictTracking,
      writesAllowed,
    ],
  )

  const cancelEditor = useCallback(() => {
    const nextEditorValue = toEditorValue(getLiveSelectedCell())
    editorValueRef.current = nextEditorValue
    setEditorValue(nextEditorValue)
    setEditorSelectionBehavior('select-all')
    editorTargetRef.current = selectionRef.current
    editingModeRef.current = 'idle'
    setEditingMode('idle')
    resetEditorConflictTracking()
  }, [getLiveSelectedCell, resetEditorConflictTracking])

  const { clearSelectedCell, copySelectionRange, fillSelectionRange, moveSelectionRange, pasteIntoSelection, toggleBooleanCell } =
    useWorkbookSelectionActions({
      writesAllowed,
      selectionRangeRef,
      selectionRef,
      editorTargetRef,
      editorValueRef,
      editingModeRef,
      invokeMutation,
      applyParsedInput,
      onPasteApplied: () => {
        perfSession.markFirstPasteApplied?.()
      },
      resetEditorConflictTracking,
      reportRuntimeError,
      setEditorValue,
      setEditingMode,
      setEditorSelectionBehavior,
    })

  const selectAddress = useCallback(
    (sheetName: string, address: string) => {
      const previousSelection = selectionRef.current
      const previousRange = selectionRangeRef.current
      if (
        editingModeRef.current === 'idle' &&
        previousSelection.sheetName === sheetName &&
        previousSelection.address === address &&
        previousRange.sheetName === sheetName &&
        previousRange.startAddress === address &&
        previousRange.endAddress === address
      ) {
        return
      }
      if (editingModeRef.current !== 'idle') {
        editorTargetRef.current = { sheetName, address }
        editingModeRef.current = 'idle'
        setEditingMode('idle')
      }
      const nextSelection = { sheetName, address }
      selectionRangeRef.current = {
        sheetName,
        startAddress: address,
        endAddress: address,
      }
      agentSelectionRangeRef.current = singleCellAgentSelectionRange(nextSelection)
      if (previousSelection.sheetName !== sheetName) {
        visibleViewportRef.current = selectionViewport(nextSelection)
      }
      selectionRef.current = nextSelection
      editorTargetRef.current = nextSelection
      resetEditorConflictTracking(nextSelection)
      runtimeActorRef.send({ type: 'selection.changed', selection: nextSelection })
    },
    [resetEditorConflictTracking, runtimeActorRef],
  )

  const handleEditorChange = useCallback(
    (next: string) => {
      if (editingModeRef.current === 'idle') {
        editorBaseSnapshotRef.current = cloneLiveSelectedCell(selectionRef.current)
        setEditorConflict(null)
      }
      editorValueRef.current = next
      setEditorValue(next)
      if (editingModeRef.current === 'idle') {
        editingModeRef.current = 'cell'
        setEditingMode('cell')
      }
    },
    [cloneLiveSelectedCell],
  )

  const isEditing = editingMode !== 'idle'
  const isEditingCell = editingMode === 'cell'
  const visibleEditorValue = isEditing ? editorValue : toEditorValue(selectedCell)
  const resolvedValue = toResolvedValue(selectedCell)
  const editorConflictBanner = useWorkbookEditorConflict({
    editingMode,
    editorValue,
    editorConflict,
    setEditorConflict,
    selectedCell,
    selection,
    editorValueRef,
    editorTargetRef,
    editorBaseSnapshotRef,
    editingModeRef,
    cloneLiveSelectedCell,
    completeEditNavigation,
    finishEditingWithAuthoritative,
    resetEditorConflictTracking,
    applyParsedInput,
    reportRuntimeError,
    setEditorSelectionBehavior,
    setEditingMode,
  })
  const handleVisibleViewportChange = useCallback((viewport: Viewport) => {
    visibleViewportRef.current = viewport
  }, [])
  const sheetNames = useMemo(
    () => [...(runtimeState?.sheetNames ?? [selection.sheetName])],
    [runtimeState?.sheetNames, selection.sheetName],
  )
  const definedNames = useMemo(() => [...(runtimeState?.definedNames ?? [])], [runtimeState])
  const getAgentContext = useCallback(
    () =>
      buildWorkbookAgentContext({
        selection: selectionRef.current,
        selectionRange: agentSelectionRangeRef.current,
        viewport: visibleViewportRef.current,
      }),
    [],
  )
  const handleSelectionRangeChange = useCallback((range: { startAddress: string; endAddress: string }) => {
    selectionRangeRef.current = {
      sheetName: selectionRef.current.sheetName,
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    }
    agentSelectionRangeRef.current = {
      startAddress: range.startAddress,
      endAddress: range.endAddress,
    }
  }, [])

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
    remoteSyncAvailable,
    changeCount,
    changesPanel,
    selectAddress,
    getAgentContext,
    previewAgentCommandBundle,
  })

  const { ribbon, statusModeLabel } = useWorkbookToolbar({
    connectionStateName: connectionState.name,
    runtimeReady,
    localPersistenceMode,
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
      void invokeColumnVisibilityMutation(selection.sheetName, selectedPosition.col, true).catch(reportRuntimeError)
    },
    onHideCurrentRow: () => {
      void invokeRowVisibilityMutation(selection.sheetName, selectedPosition.row, true).catch(reportRuntimeError)
    },
    onRedo: redoLatestChange,
    onUndo: undoLatestChange,
    onUnhideCurrentColumn: () => {
      void invokeColumnVisibilityMutation(selection.sheetName, selectedPosition.col, false).catch(reportRuntimeError)
    },
    onUnhideCurrentRow: () => {
      void invokeRowVisibilityMutation(selection.sheetName, selectedPosition.row, false).catch(reportRuntimeError)
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

  const subscribeViewport = useCallback(
    (
      sheetName: string,
      viewport: Parameters<ProjectedViewportStore['subscribeViewport']>[1],
      listener: Parameters<ProjectedViewportStore['subscribeViewport']>[2],
    ) => {
      if (!runtimeController) {
        return () => {}
      }
      return runtimeController.subscribeViewport(sheetName, viewport, listener)
    },
    [runtimeController],
  )

  const { createSheet, deleteSheet, renameSheet } = useWorkbookSheetActions({
    sheetNames,
    selectionRef,
    invokeMutation,
    selectAddress,
    reportRuntimeError,
  })

  const invokeInsertRowsMutation = useCallback(
    async (sheetName: string, startRow: number, count: number): Promise<void> => {
      await invokeMutation('insertRows', sheetName, startRow, count)
    },
    [invokeMutation],
  )

  const invokeDeleteRowsMutation = useCallback(
    async (sheetName: string, startRow: number, count: number): Promise<void> => {
      await invokeMutation('deleteRows', sheetName, startRow, count)
    },
    [invokeMutation],
  )

  const invokeInsertColumnsMutation = useCallback(
    async (sheetName: string, startCol: number, count: number): Promise<void> => {
      await invokeMutation('insertColumns', sheetName, startCol, count)
    },
    [invokeMutation],
  )

  const invokeDeleteColumnsMutation = useCallback(
    async (sheetName: string, startCol: number, count: number): Promise<void> => {
      await invokeMutation('deleteColumns', sheetName, startCol, count)
    },
    [invokeMutation],
  )

  const invokeSetFreezePaneMutation = useCallback(
    async (sheetName: string, rows: number, cols: number): Promise<void> => {
      await invokeMutation('setFreezePane', sheetName, rows, cols)
    },
    [invokeMutation],
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
    approvePersistenceTransfer,
    selectAddress,
    selectedCell,
    selection,
    sidePanelId,
    dismissPersistenceTransferRequest,
    handleSelectionRangeChange,
    setSidePanelWidth,
    sheetNames,
    sidePanel,
    sidePanelWidth,
    statusModeLabel,
    localPersistenceMode,
    pendingTransferRequest,
    requestPersistenceTransfer,
    subscribeViewport,
    transferRequested,
    toggleBooleanCell,
    visibleEditorValue,
    workbookReady,
    workerHandle,
    writesAllowed,
    zeroConfigured,
  }
}
