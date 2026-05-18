import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject } from 'react'
import type { EditMovement, EditSelectionBehavior, GridSelectionSnapshot } from '@bilig/grid'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import { scheduleSelectionPersistence } from './selection-persistence.js'
import type { WorkbookPerfSession } from './perf/workbook-perf.js'
import type { WorkerHandle, WorkerRuntimeSelection } from './runtime-session.js'
import { useWorkbookEditorConflict } from './use-workbook-editor-conflict.js'
import { useWorkbookSelectionActions } from './use-workbook-selection-actions.js'
import { createSingleCellSelectionSnapshot } from './workbook-agent-context.js'
import {
  createOptimisticCellSnapshot,
  createSupersedingCellSnapshot,
  evaluateOptimisticFormula,
  optimisticCellKey,
} from './workbook-optimistic-cell.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import {
  clampSelectionMovement,
  emptyCellSnapshot,
  parseEditorInput,
  parsedEditorInputEquals,
  parsedEditorInputFromSnapshot,
  parsedEditorInputMatchesSnapshot,
  sameCellContent,
  toEditorValue,
  toResolvedValue,
  type EditingMode,
  type ParsedEditorInput,
  type WorkbookEditorConflict,
} from './worker-workbook-app-model.js'

export interface EditTargetSelection {
  readonly sheetName: string
  readonly address: string
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

function resolveDetachedOptimisticValue(
  targetSelection: EditTargetSelection,
  selectedCell: CellSnapshot,
  parsed: ParsedEditorInput,
): string {
  const current =
    selectedCell.sheetName === targetSelection.sheetName && selectedCell.address === targetSelection.address
      ? selectedCell
      : emptyCellSnapshot(targetSelection.sheetName, targetSelection.address)
  return toResolvedValue(
    createOptimisticCellSnapshot({
      sheetName: targetSelection.sheetName,
      address: targetSelection.address,
      current,
      parsed,
    }),
  )
}

export function useWorkerWorkbookInteractionState(input: {
  documentId: string
  currentUserId: string
  selection: WorkerRuntimeSelection
  selectedCell: CellSnapshot
  workerHandle: WorkerHandle | null | undefined
  workerHandleRef: MutableRefObject<WorkerHandle | null>
  writesAllowed: boolean
  invokeMutation: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>
  perfSession: WorkbookPerfSession
  reportRuntimeError: (error: unknown) => void
  sendSelectionChanged: (selection: WorkerRuntimeSelection) => void
  onSelectionSheetChanged?: (nextSelection: WorkerRuntimeSelection, previousSelection: WorkerRuntimeSelection) => void
}) {
  const {
    currentUserId,
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
    onSelectionSheetChanged,
  } = input

  const [editorValue, setEditorValue] = useState('')
  const [editorSelectionBehavior, setEditorSelectionBehavior] = useState<EditSelectionBehavior>('select-all')
  const [editingMode, setEditingMode] = useState<EditingMode>('idle')
  const [editorConflict, setEditorConflict] = useState<WorkbookEditorConflict | null>(null)
  const [selectionSnapshot, setSelectionSnapshot] = useState<GridSelectionSnapshot>(createSingleCellSelectionSnapshot(selection))
  const [, bumpOptimisticSeedRevision] = useState(0)

  const selectionRef = useRef(selection)
  const editorValueRef = useRef(editorValue)
  const editingModeRef = useRef(editingMode)
  const editorTargetRef = useRef(selection)
  const editorBaseSnapshotRef = useRef<CellSnapshot>(emptyCellSnapshot(selection.sheetName, selection.address))
  const selectionSnapshotRef = useRef<GridSelectionSnapshot>(createSingleCellSelectionSnapshot(selection))
  const selectionRangeRef = useRef<CellRangeRef>(selectionSnapshotToRangeRef(selectionSnapshotRef.current))
  const pendingExternalSelectionRef = useRef<GridSelectionSnapshot | null>(null)
  const pendingLocalSelectionRef = useRef<GridSelectionSnapshot | null>(null)
  const optimisticCellSeedsRef = useRef<Map<string, string>>(new Map())
  const optimisticCellResolvedValuesRef = useRef<Map<string, string>>(new Map())
  const editSessionRef = useRef(0)
  const pendingEditCommitSessionRef = useRef<number | null>(null)
  const pendingEditCommitMovementAppliedRef = useRef(false)

  useEffect(() => {
    const previousSelection = selectionRef.current
    const activeExternalSelection = pendingExternalSelectionRef.current
    const activeLocalSelection = pendingLocalSelectionRef.current
    if (
      activeExternalSelection &&
      (activeExternalSelection.sheetName !== selection.sheetName || activeExternalSelection.address !== selection.address)
    ) {
      return
    }
    if (
      activeLocalSelection &&
      (activeLocalSelection.sheetName !== selection.sheetName || activeLocalSelection.address !== selection.address)
    ) {
      return
    }
    selectionRef.current = selection
    if (
      activeLocalSelection &&
      activeLocalSelection.sheetName === selection.sheetName &&
      activeLocalSelection.address === selection.address
    ) {
      pendingLocalSelectionRef.current = null
    }
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
    editorValueRef.current = editorValue
  }, [editorValue])

  useEffect(() => {
    editingModeRef.current = editingMode
  }, [editingMode])

  const getLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current
      if (!active) {
        return selectedCell
      }
      return active.viewportStore.getCell(nextSelection.sheetName, nextSelection.address)
    },
    [selectedCell, workerHandleRef],
  )

  const getCellEditorSeed = useCallback((sheetName: string, address: string) => {
    return optimisticCellSeedsRef.current.get(optimisticCellKey(sheetName, address))
  }, [])

  const clearOptimisticCellSeed = useCallback((sheetName: string, address: string, seed: string) => {
    const key = optimisticCellKey(sheetName, address)
    if (optimisticCellSeedsRef.current.get(key) === seed) {
      optimisticCellSeedsRef.current.delete(key)
      optimisticCellResolvedValuesRef.current.delete(key)
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])
  const supersedeOptimisticCellSeedsForRange = useCallback((range: CellRangeRef): (() => void) | null => {
    const start = parseCellAddress(range.startAddress, range.sheetName)
    const end = parseCellAddress(range.endAddress, range.sheetName)
    const startRow = Math.min(start.row, end.row)
    const endRow = Math.max(start.row, end.row)
    const startCol = Math.min(start.col, end.col)
    const endCol = Math.max(start.col, end.col)
    const removedSeeds: Array<readonly [string, string, string | undefined]> = []

    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        const address = formatAddress(row, col)
        const key = optimisticCellKey(range.sheetName, address)
        const seed = optimisticCellSeedsRef.current.get(key)
        if (seed === undefined) {
          continue
        }
        removedSeeds.push([key, seed, optimisticCellResolvedValuesRef.current.get(key)])
        optimisticCellSeedsRef.current.delete(key)
        optimisticCellResolvedValuesRef.current.delete(key)
      }
    }

    if (removedSeeds.length === 0) {
      return null
    }

    bumpOptimisticSeedRevision((revision) => revision + 1)
    return () => {
      for (const [key, seed, resolvedValue] of removedSeeds) {
        if (!optimisticCellSeedsRef.current.has(key)) {
          optimisticCellSeedsRef.current.set(key, seed)
        }
        if (resolvedValue !== undefined && !optimisticCellResolvedValuesRef.current.has(key)) {
          optimisticCellResolvedValuesRef.current.set(key, resolvedValue)
        }
      }
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])
  const supersedeOptimisticCellSeedsForSheet = useCallback((sheetName: string): (() => void) | null => {
    const prefix = `${sheetName}:`
    const removedSeeds: Array<readonly [string, string, string | undefined]> = []

    for (const [key, seed] of optimisticCellSeedsRef.current) {
      if (key.startsWith(prefix)) {
        removedSeeds.push([key, seed, optimisticCellResolvedValuesRef.current.get(key)])
      }
    }

    if (removedSeeds.length === 0) {
      return null
    }

    for (const [key] of removedSeeds) {
      optimisticCellSeedsRef.current.delete(key)
      optimisticCellResolvedValuesRef.current.delete(key)
    }

    bumpOptimisticSeedRevision((revision) => revision + 1)
    return () => {
      for (const [key, seed, resolvedValue] of removedSeeds) {
        if (!optimisticCellSeedsRef.current.has(key)) {
          optimisticCellSeedsRef.current.set(key, seed)
        }
        if (resolvedValue !== undefined && !optimisticCellResolvedValuesRef.current.has(key)) {
          optimisticCellResolvedValuesRef.current.set(key, resolvedValue)
        }
      }
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])
  const replaceOptimisticCellSeed = useCallback((sheetName: string, address: string, seed: string): (() => void) => {
    const key = optimisticCellKey(sheetName, address)
    const hadPreviousSeed = optimisticCellSeedsRef.current.has(key)
    const previousSeed = optimisticCellSeedsRef.current.get(key)
    optimisticCellSeedsRef.current.set(key, seed)
    bumpOptimisticSeedRevision((revision) => revision + 1)

    return () => {
      if (hadPreviousSeed && previousSeed !== undefined) {
        optimisticCellSeedsRef.current.set(key, previousSeed)
      } else {
        optimisticCellSeedsRef.current.delete(key)
      }
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])

  const applyOptimisticParsedInput = useCallback(
    (targetSelection: EditTargetSelection, parsed: ParsedEditorInput) => {
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
      return {
        editorSeed: toEditorValue(optimistic),
        resolvedValue: toResolvedValue(optimistic),
        rollback: (snapshot = previous) => {
          viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, optimistic.version + 1))
        },
      }
    },
    [workerHandleRef],
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

  const completeSelectionNavigation = useCallback(
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
      const nextSelectionSnapshot = createSingleCellSelectionSnapshot(nextSelection)
      selectionRef.current = nextSelection
      editorTargetRef.current = nextSelection
      selectionSnapshotRef.current = nextSelectionSnapshot
      selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
      setSelectionSnapshot(nextSelectionSnapshot)
      pendingLocalSelectionRef.current = nextSelectionSnapshot
      sendSelectionChanged(nextSelection)
      return nextSelection
    },
    [sendSelectionChanged],
  )

  const finishEditingAtSelection = useCallback(
    (nextSelection: WorkerRuntimeSelection) => {
      const nextEditorValue = toEditorValue(cloneLiveSelectedCell(nextSelection))
      editorValueRef.current = nextEditorValue
      setEditorValue(nextEditorValue)
      setEditorSelectionBehavior('select-all')
      editingModeRef.current = 'idle'
      setEditingMode('idle')
      resetEditorConflictTracking(nextSelection)
    },
    [cloneLiveSelectedCell, resetEditorConflictTracking],
  )

  const finishEditingWithAuthoritative = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      finishEditingAtSelection(completeSelectionNavigation(targetSelection, movement))
    },
    [completeSelectionNavigation, finishEditingAtSelection],
  )

  const beginEditing = useCallback(
    (
      seed?: string,
      selectionBehavior: EditSelectionBehavior = 'select-all',
      mode: Exclude<EditingMode, 'idle'> = 'cell',
      targetSelectionOverride?: EditTargetSelection,
    ) => {
      if (!writesAllowed) {
        return
      }
      editSessionRef.current += 1
      pendingEditCommitSessionRef.current = null
      pendingEditCommitMovementAppliedRef.current = false
      const previousSelection = selectionRef.current
      const nextTarget = targetSelectionOverride ?? previousSelection
      const targetChanged = previousSelection.sheetName !== nextTarget.sheetName || previousSelection.address !== nextTarget.address
      const nextSelectionSnapshot = createSingleCellSelectionSnapshot(nextTarget)
      selectionSnapshotRef.current = nextSelectionSnapshot
      selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
      setSelectionSnapshot(nextSelectionSnapshot)
      pendingExternalSelectionRef.current = null
      pendingLocalSelectionRef.current = targetChanged ? nextSelectionSnapshot : null
      selectionRef.current = nextTarget
      if (targetChanged) {
        onSelectionSheetChanged?.(nextTarget, previousSelection)
        sendSelectionChanged(nextTarget)
      }
      const nextEditorValue = seed ?? toEditorValue(getLiveSelectedCell(nextTarget))
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextTarget)
      setEditorConflict(null)
      editorValueRef.current = nextEditorValue
      setEditorValue(nextEditorValue)
      setEditorSelectionBehavior(selectionBehavior)
      editorTargetRef.current = nextTarget
      editingModeRef.current = mode
      setEditingMode(mode)
    },
    [cloneLiveSelectedCell, getLiveSelectedCell, onSelectionSheetChanged, sendSelectionChanged, writesAllowed],
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
        if (movement && !pendingEditCommitMovementAppliedRef.current) {
          pendingEditCommitMovementAppliedRef.current = true
          completeSelectionNavigation(selectionRef.current, movement)
        }
        return
      }
      const targetSelection =
        targetSelectionOverride ?? (editingModeRef.current === 'idle' ? selectionRef.current : editorTargetRef.current)
      const nextValue =
        valueOverride ??
        (editingModeRef.current === 'idle'
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : (readMountedCellEditorValue() ?? editorValueRef.current))
      const commitSessionId = editSessionRef.current
      if (pendingEditCommitSessionRef.current === commitSessionId) {
        if (movement && !pendingEditCommitMovementAppliedRef.current) {
          pendingEditCommitMovementAppliedRef.current = true
          finishEditingWithAuthoritative(targetSelection, movement)
        }
        return
      }
      pendingEditCommitSessionRef.current = commitSessionId
      pendingEditCommitMovementAppliedRef.current = false
      const parsed = parseEditorInput(nextValue)
      const isFreshValueOverride = valueOverride !== undefined && editingModeRef.current === 'idle'
      const baseSnapshot = editorBaseSnapshotRef.current
      const liveSnapshot = cloneLiveSelectedCell(targetSelection)
      const baseMatchesTarget = baseSnapshot.sheetName === targetSelection.sheetName && baseSnapshot.address === targetSelection.address
      const targetBaseSnapshot = isFreshValueOverride ? liveSnapshot : baseMatchesTarget ? baseSnapshot : liveSnapshot
      const draftMatchesLiveSnapshot = parsedEditorInputMatchesSnapshot(parsed, liveSnapshot)
      const draftMatchesBase = parsedEditorInputEquals(parsed, parsedEditorInputFromSnapshot(targetBaseSnapshot))

      if (!sameCellContent(targetBaseSnapshot, liveSnapshot) && !draftMatchesLiveSnapshot) {
        pendingEditCommitSessionRef.current = null
        pendingEditCommitMovementAppliedRef.current = false
        if (draftMatchesBase) {
          pendingEditCommitMovementAppliedRef.current = Boolean(movement)
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
        pendingEditCommitMovementAppliedRef.current = Boolean(movement)
        finishEditingWithAuthoritative(targetSelection, movement)
        return
      }

      const nextSelection = completeSelectionNavigation(targetSelection, movement)
      pendingEditCommitMovementAppliedRef.current = Boolean(movement)
      const optimisticResult = applyOptimisticParsedInput(targetSelection, parsed)
      const optimisticEditorSeed = optimisticResult?.editorSeed ?? nextValue
      const optimisticKey = optimisticCellKey(targetSelection.sheetName, targetSelection.address)
      optimisticCellSeedsRef.current.set(optimisticKey, optimisticEditorSeed)
      optimisticCellResolvedValuesRef.current.set(
        optimisticKey,
        optimisticResult?.resolvedValue ?? resolveDetachedOptimisticValue(targetSelection, selectedCell, parsed),
      )
      bumpOptimisticSeedRevision((revision) => revision + 1)
      const mutationTask = applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed)
      finishEditingAtSelection(nextSelection)
      void (async () => {
        try {
          await mutationTask
          if (editSessionRef.current !== commitSessionId) {
            return
          }
        } catch (error) {
          clearOptimisticCellSeed(targetSelection.sheetName, targetSelection.address, optimisticEditorSeed)
          optimisticResult?.rollback()
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
      applyOptimisticParsedInput,
      applyParsedInput,
      clearOptimisticCellSeed,
      cloneLiveSelectedCell,
      completeSelectionNavigation,
      finishEditingAtSelection,
      finishEditingWithAuthoritative,
      getLiveSelectedCell,
      reportRuntimeError,
      selectedCell,
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
      supersedeOptimisticCellSeedsForRange,
      replaceOptimisticCellSeed,
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
      if (wasEditing) {
        commitEditor(undefined, undefined, editorTargetRef.current)
        if (editingModeRef.current !== 'idle') {
          return
        }
      }
      if (
        previousSelection.sheetName === sheetName &&
        previousSelection.address === address &&
        previousRange.sheetName === sheetName &&
        previousRange.startAddress === nextSelectionSnapshot.range.startAddress &&
        previousRange.endAddress === nextSelectionSnapshot.range.endAddress
      ) {
        if (options?.markAsExternal) {
          sendSelectionChanged({ sheetName, address })
        }
        return
      }
      const nextSelection = { sheetName, address }
      selectionSnapshotRef.current = nextSelectionSnapshot
      selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
      setSelectionSnapshot(nextSelectionSnapshot)
      pendingExternalSelectionRef.current = options?.markAsExternal ? nextSelectionSnapshot : null
      pendingLocalSelectionRef.current = options?.markAsExternal ? null : nextSelectionSnapshot
      if (previousSelection.sheetName !== sheetName) {
        onSelectionSheetChanged?.(nextSelection, previousSelection)
      }
      selectionRef.current = nextSelection
      editorTargetRef.current = nextSelection
      setEditorConflict((current) => (current === null ? current : null))
      sendSelectionChanged(nextSelection)
    },
    [commitEditor, onSelectionSheetChanged, sendSelectionChanged],
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

  const handleSelectionChange = useCallback(
    (nextSelectionSnapshot: GridSelectionSnapshot) => {
      const activeExternalSelection = pendingExternalSelectionRef.current
      if (activeExternalSelection && !selectionSnapshotsEqual(activeExternalSelection, nextSelectionSnapshot)) {
        pendingExternalSelectionRef.current = null
      }
      applySelectionSnapshot(nextSelectionSnapshot, { markAsExternal: false })
    },
    [applySelectionSnapshot],
  )
  const acknowledgeExternalSelectionSync = useCallback((syncedSelectionSnapshot: GridSelectionSnapshot) => {
    const activeExternalSelection = pendingExternalSelectionRef.current
    if (!activeExternalSelection || !selectionSnapshotsEqual(activeExternalSelection, syncedSelectionSnapshot)) {
      return
    }
    pendingExternalSelectionRef.current = null
  }, [])

  const handleEditorChange = useCallback(
    (next: string) => {
      if (editingModeRef.current === 'idle') {
        const nextTarget = selectionRef.current
        editSessionRef.current += 1
        pendingEditCommitSessionRef.current = null
        pendingEditCommitMovementAppliedRef.current = false
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
  const visibleSelection = useMemo(
    () => ({
      sheetName: selectionSnapshot.sheetName,
      address: selectionSnapshot.address,
    }),
    [selectionSnapshot.address, selectionSnapshot.sheetName],
  )

  useEffect(() => {
    scheduleSelectionPersistence({ documentId, userId: currentUserId }, visibleSelection)
  }, [currentUserId, documentId, visibleSelection])

  const visibleCellKey = optimisticCellKey(visibleSelection.sheetName, visibleSelection.address)
  const liveVisibleCell = getLiveSelectedCell(visibleSelection)
  const visibleEditorValue = isEditing
    ? editorValue
    : (getCellEditorSeed(visibleSelection.sheetName, visibleSelection.address) ?? toEditorValue(liveVisibleCell))
  const visibleResolvedValue = optimisticCellResolvedValuesRef.current.get(visibleCellKey) ?? toResolvedValue(liveVisibleCell)

  useEffect(() => {
    const optimisticSeed = optimisticCellSeedsRef.current.get(visibleCellKey)
    if (optimisticSeed === undefined) {
      return
    }
    if (optimisticSeed === toEditorValue(liveVisibleCell)) {
      optimisticCellSeedsRef.current.delete(visibleCellKey)
      optimisticCellResolvedValuesRef.current.delete(visibleCellKey)
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [liveVisibleCell, visibleCellKey])

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
    completeEditNavigation: completeSelectionNavigation,
    finishEditingWithAuthoritative,
    resetEditorConflictTracking,
    applyParsedInput,
    reportRuntimeError,
    setEditorSelectionBehavior,
    setEditingMode,
  })

  return {
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
    selectedCell,
    selectionRef,
    selectionRangeRef,
    selectionSnapshot,
    selectionSnapshotRef,
    selectSelectionSnapshot,
    supersedeOptimisticCellSeedsForSheet,
    toggleBooleanCell,
    visibleSelectedCell: liveVisibleCell,
    visibleEditorValue,
    visibleResolvedValue,
    visibleSelection,
  }
}
