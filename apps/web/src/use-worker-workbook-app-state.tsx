import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore, type ReactNode } from 'react'
import { useActorRef, useSelector } from '@xstate/react'
import { isWorkbookAgentCommandBundle, isWorkbookAgentPreviewSummary, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { EditMovement, EditSelectionBehavior, GridSelectionSnapshot } from '@bilig/grid'
import { parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, CellSnapshot, Viewport } from '@bilig/protocol'
import { createWorkerRuntimeMachine } from './runtime-machine.js'
import type { resolveRuntimeConfig } from './runtime-config.js'
import type { WorkerRuntimeSelection, ZeroClient } from './runtime-session.js'
import { loadPersistedSelection, persistSelection } from './selection-persistence.js'
import type { ProjectedViewportStore } from './projected-viewport-store.js'
import { buildWorkbookAgentContext, createSingleCellSelectionSnapshot } from './workbook-agent-context.js'
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
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import { registerRuntimeDisposalHandlers } from './runtime-disposal-handlers.js'
import { useWorkbookLocalPersistenceHandoff } from './use-workbook-local-persistence-handoff.js'
import { loadOrCreateWorkbookPresenceClientId } from './workbook-presence-client.js'
import {
  createOptimisticCellSnapshot,
  createSupersedingCellSnapshot,
  evaluateOptimisticFormula,
  optimisticCellKey,
} from './workbook-optimistic-cell.js'

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

function selectionSnapshotToRangeRef(selection: GridSelectionSnapshot): CellRangeRef {
  return {
    sheetName: selection.sheetName,
    startAddress: selection.range.startAddress,
    endAddress: selection.range.endAddress,
  }
}

function selectionSnapshotsEqual(left: GridSelectionSnapshot, right: GridSelectionSnapshot): boolean {
  return (
    left.sheetName === right.sheetName &&
    left.address === right.address &&
    left.kind === right.kind &&
    left.range.startAddress === right.range.startAddress &&
    left.range.endAddress === right.range.endAddress
  )
}

function readMountedCellEditorValue(): string | null {
  if (typeof document === 'undefined') {
    return null
  }
  const editor = document.querySelector<HTMLTextAreaElement>('[data-testid="cell-editor-input"]')
  return editor?.value ?? null
}

interface EditTargetSelection {
  readonly sheetName: string
  readonly address: string
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
  const [selectionSnapshot, setSelectionSnapshot] = useState<GridSelectionSnapshot>(createSingleCellSelectionSnapshot(selection))
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
  const selectionSnapshotRef = useRef<GridSelectionSnapshot>(createSingleCellSelectionSnapshot(selection))
  const selectionRangeRef = useRef<CellRangeRef>(selectionSnapshotToRangeRef(selectionSnapshotRef.current))
  const pendingExternalSelectionRef = useRef<GridSelectionSnapshot | null>(null)
  const optimisticCellSeedsRef = useRef<Map<string, string>>(new Map())
  const editSessionRef = useRef(0)
  const pendingEditCommitSessionRef = useRef<number | null>(null)

  useEffect(() => {
    const previousSelection = selectionRef.current
    selectionRef.current = selection
    const activeExternalSelection = pendingExternalSelectionRef.current
    if (
      activeExternalSelection &&
      activeExternalSelection.sheetName === selection.sheetName &&
      activeExternalSelection.address === selection.address
    ) {
      pendingExternalSelectionRef.current = null
    }
    if (previousSelection.sheetName !== selection.sheetName || previousSelection.address !== selection.address) {
      const activeSelectionSnapshot = selectionSnapshotRef.current
      if (activeSelectionSnapshot.sheetName === selection.sheetName && activeSelectionSnapshot.address === selection.address) {
        return
      }
      const nextSelectionSnapshot = createSingleCellSelectionSnapshot(selection)
      selectionSnapshotRef.current = nextSelectionSnapshot
      selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
      setSelectionSnapshot(nextSelectionSnapshot)
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
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'columnWidths', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
    () => readViewportColumnWidths(workerHandle, selection.sheetName),
  )

  const rowHeights = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'rowHeights', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
    () => readViewportRowHeights(workerHandle, selection.sheetName),
  )

  const hiddenColumns = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenColumns', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
    () => readViewportHiddenColumns(workerHandle, selection.sheetName),
  )

  const hiddenRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) =>
        workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'hiddenRows', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
    () => readViewportHiddenRows(workerHandle, selection.sheetName),
  )

  const freezeRows = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeRows(selection.sheetName) ?? 0,
  )

  const freezeCols = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeSheetChannel(selection.sheetName, 'freeze', listener) ?? (() => {}),
      [selection.sheetName, workerHandle],
    ),
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
    () => workerHandle?.viewportStore.getFreezeCols(selection.sheetName) ?? 0,
  )

  const selectedCell = useSyncExternalStore(
    useCallback(
      (listener: () => void) => workerHandle?.viewportStore.subscribeCell(selection.sheetName, selection.address, listener) ?? (() => {}),
      [selection.address, selection.sheetName, workerHandle],
    ),
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
  const getCellEditorSeed = useCallback((sheetName: string, address: string) => {
    return optimisticCellSeedsRef.current.get(optimisticCellKey(sheetName, address))
  }, [])
  const clearOptimisticCellSeed = useCallback((sheetName: string, address: string, seed: string) => {
    const key = optimisticCellKey(sheetName, address)
    if (optimisticCellSeedsRef.current.get(key) === seed) {
      optimisticCellSeedsRef.current.delete(key)
    }
  }, [])
  const applyOptimisticParsedInput = useCallback((targetSelection: EditTargetSelection, parsed: ParsedEditorInput) => {
    const viewportStore = workerHandleRef.current?.viewportStore
    if (!viewportStore) {
      return null
    }
    const previous = viewportStore.getCell(targetSelection.sheetName, targetSelection.address)
    const optimistic = createOptimisticCellSnapshot({
      sheetName: targetSelection.sheetName,
      address: targetSelection.address,
      current: previous,
      parsed,
      evaluateFormula: (formula) =>
        evaluateOptimisticFormula({
          sheetName: targetSelection.sheetName,
          address: targetSelection.address,
          formula,
          getCell: (sheetName, address) => viewportStore.getCell(sheetName, address),
        }),
    })
    viewportStore.setCellSnapshot(optimistic)
    return (snapshot = previous) => {
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, optimistic.version + 1))
    }
  }, [])

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
      const syncSelectionSnapshot = (nextSelection: WorkerRuntimeSelection) => {
        const nextSelectionSnapshot = createSingleCellSelectionSnapshot(nextSelection)
        selectionSnapshotRef.current = nextSelectionSnapshot
        selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
        setSelectionSnapshot(nextSelectionSnapshot)
      }
      if (!movement) {
        selectionRef.current = targetSelection
        editorTargetRef.current = targetSelection
        syncSelectionSnapshot(targetSelection)
        return targetSelection
      }
      const nextAddress = clampSelectionMovement(targetSelection.address, targetSelection.sheetName, movement)
      const nextSelection = { sheetName: targetSelection.sheetName, address: nextAddress }
      selectionRef.current = nextSelection
      editorTargetRef.current = nextSelection
      syncSelectionSnapshot(nextSelection)
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
      editSessionRef.current += 1
      pendingEditCommitSessionRef.current = null
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
    (movement?: EditMovement, valueOverride?: string, targetSelectionOverride?: EditTargetSelection) => {
      if (!writesAllowed) {
        return
      }
      if (editingModeRef.current === 'idle' && valueOverride === undefined && targetSelectionOverride === undefined) {
        return
      }
      const targetSelection =
        targetSelectionOverride ?? (editingModeRef.current === 'idle' ? selectionRef.current : editorTargetRef.current)
      const nextValue =
        editingModeRef.current === 'idle'
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : (valueOverride ?? readMountedCellEditorValue() ?? editorValueRef.current)
      const commitSessionId = editSessionRef.current
      if (pendingEditCommitSessionRef.current === commitSessionId) {
        return
      }
      pendingEditCommitSessionRef.current = commitSessionId
      const parsed = parseEditorInput(nextValue)
      const baseSnapshot = editorBaseSnapshotRef.current
      const liveSnapshot = cloneLiveSelectedCell(targetSelection)
      const targetBaseSnapshot =
        baseSnapshot.sheetName === targetSelection.sheetName && baseSnapshot.address === targetSelection.address
          ? baseSnapshot
          : liveSnapshot
      const draftMatchesLiveSnapshot = parsedEditorInputMatchesSnapshot(parsed, liveSnapshot)
      const draftMatchesBase = parsedEditorInputEquals(parsed, parsedEditorInputFromSnapshot(targetBaseSnapshot))

      if (!sameCellContent(targetBaseSnapshot, liveSnapshot) && !draftMatchesLiveSnapshot) {
        pendingEditCommitSessionRef.current = null
        if (draftMatchesBase) {
          finishEditingWithAuthoritative(targetSelection, movement)
          return
        }
        setEditorConflict({
          sheetName: targetSelection.sheetName,
          address: targetSelection.address,
          phase: 'compare',
          baseSnapshot: targetBaseSnapshot,
          authoritativeSnapshot: liveSnapshot,
        })
        return
      }

      if (draftMatchesLiveSnapshot) {
        pendingEditCommitSessionRef.current = null
        finishEditingWithAuthoritative(targetSelection, movement)
        return
      }

      const nextSelection = completeEditNavigation(targetSelection, movement)
      optimisticCellSeedsRef.current.set(optimisticCellKey(targetSelection.sheetName, targetSelection.address), nextValue)
      const rollbackOptimisticCell = applyOptimisticParsedInput(targetSelection, parsed)
      void (async () => {
        try {
          await applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed)
          clearOptimisticCellSeed(targetSelection.sheetName, targetSelection.address, nextValue)
          if (editSessionRef.current !== commitSessionId) {
            return
          }
          setEditorSelectionBehavior('select-all')
          editingModeRef.current = 'idle'
          setEditingMode('idle')
          resetEditorConflictTracking(nextSelection)
        } catch (error) {
          clearOptimisticCellSeed(targetSelection.sheetName, targetSelection.address, nextValue)
          rollbackOptimisticCell?.()
          if (editSessionRef.current !== commitSessionId) {
            return
          }
          editingModeRef.current = 'idle'
          setEditingMode('idle')
          reportRuntimeError(error)
        } finally {
          if (pendingEditCommitSessionRef.current === commitSessionId) {
            pendingEditCommitSessionRef.current = null
          }
        }
      })()
    },
    [
      applyParsedInput,
      cloneLiveSelectedCell,
      completeEditNavigation,
      finishEditingWithAuthoritative,
      getLiveSelectedCell,
      clearOptimisticCellSeed,
      applyOptimisticParsedInput,
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
      viewportStore: workerHandle?.viewportStore,
      onPasteApplied: () => {
        perfSession.markFirstPasteApplied?.()
      },
      resetEditorConflictTracking,
      reportRuntimeError,
      setEditorValue,
      setEditingMode,
      setEditorSelectionBehavior,
    })

  const applySelectionSnapshot = useCallback(
    (nextSelectionSnapshot: GridSelectionSnapshot, options?: { markAsExternal?: boolean }) => {
      const { sheetName, address } = nextSelectionSnapshot
      const previousSelection = selectionRef.current
      const previousRange = selectionRangeRef.current
      const wasEditing = editingModeRef.current !== 'idle'
      if (
        !wasEditing &&
        previousSelection.sheetName === sheetName &&
        previousSelection.address === address &&
        previousRange.sheetName === sheetName &&
        previousRange.startAddress === nextSelectionSnapshot.range.startAddress &&
        previousRange.endAddress === nextSelectionSnapshot.range.endAddress
      ) {
        return
      }
      if (wasEditing) {
        editingModeRef.current = 'idle'
        setEditingMode('idle')
      }
      const nextSelection = { sheetName, address }
      selectionSnapshotRef.current = nextSelectionSnapshot
      selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
      setSelectionSnapshot(nextSelectionSnapshot)
      pendingExternalSelectionRef.current = options?.markAsExternal ? nextSelectionSnapshot : null
      if (previousSelection.sheetName !== sheetName) {
        visibleViewportRef.current = selectionViewport(nextSelection)
      }
      selectionRef.current = nextSelection
      if (!wasEditing) {
        editorTargetRef.current = nextSelection
        resetEditorConflictTracking(nextSelection)
      }
      runtimeActorRef.send({ type: 'selection.changed', selection: nextSelection })
    },
    [resetEditorConflictTracking, runtimeActorRef],
  )

  const selectAddress = useCallback(
    (sheetName: string, address: string) => {
      applySelectionSnapshot(
        createSingleCellSelectionSnapshot({
          sheetName,
          address,
        }),
        { markAsExternal: true },
      )
    },
    [applySelectionSnapshot],
  )

  const selectSelectionSnapshot = useCallback(
    (nextSelectionSnapshot: GridSelectionSnapshot) => {
      applySelectionSnapshot(nextSelectionSnapshot, { markAsExternal: true })
    },
    [applySelectionSnapshot],
  )

  const handleEditorChange = useCallback(
    (next: string) => {
      if (editingModeRef.current === 'idle') {
        const nextTarget = selectionRef.current
        editSessionRef.current += 1
        pendingEditCommitSessionRef.current = null
        editorTargetRef.current = nextTarget
        editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextTarget)
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
  const visibleEditorValue = isEditing
    ? editorValue
    : (getCellEditorSeed(selection.sheetName, selection.address) ?? toEditorValue(selectedCell))
  const resolvedValue = toResolvedValue(selectedCell)
  useEffect(() => {
    const key = optimisticCellKey(selection.sheetName, selection.address)
    const optimisticSeed = optimisticCellSeedsRef.current.get(key)
    if (optimisticSeed === undefined) {
      return
    }
    if (optimisticSeed === toEditorValue(selectedCell)) {
      optimisticCellSeedsRef.current.delete(key)
    }
  }, [selectedCell, selection.address, selection.sheetName])
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
        selection: selectionSnapshotRef.current,
        viewport: visibleViewportRef.current,
      }),
    [],
  )
  const handleSelectionChange = useCallback(
    (nextSelectionSnapshot: GridSelectionSnapshot) => {
      const activeExternalSelection = pendingExternalSelectionRef.current
      if (activeExternalSelection && !selectionSnapshotsEqual(activeExternalSelection, nextSelectionSnapshot)) {
        return
      }
      applySelectionSnapshot(nextSelectionSnapshot, { markAsExternal: false })
    },
    [applySelectionSnapshot],
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
    handleSelectionChange,
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
