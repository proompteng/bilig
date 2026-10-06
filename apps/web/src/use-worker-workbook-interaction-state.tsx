import { useCallback, useEffect, useMemo, useRef, useState, type MutableRefObject } from 'react'
import type { EditMovement, EditSelectionBehavior, GridSelectionSnapshot } from '@bilig/grid'
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
import { optimisticCellTargetFromKey, supersedeOptimisticSeedsInRange } from './workbook-optimistic-seeds.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'
import { LOCAL_CELL_CONTENT_DIRTY_MASK } from './projected-workbook-local-delta.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import { useDeferredEditCommits } from './use-deferred-edit-commits.js'
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

function workerSelectionAcknowledgesSnapshot(snapshot: GridSelectionSnapshot, selection: WorkerRuntimeSelection): boolean {
  return (
    snapshot.kind === 'cell' &&
    snapshot.sheetName === selection.sheetName &&
    snapshot.address === selection.address &&
    snapshot.range.startAddress === selection.address &&
    snapshot.range.endAddress === selection.address
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

function isOptimisticSnapshot(snapshot: CellSnapshot): boolean {
  return (snapshot.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) !== 0
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
  invokeEditCommitMutation?: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>
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
    invokeEditCommitMutation = invokeMutation,
    perfSession,
    reportRuntimeError,
    sendSelectionChanged,
    onSelectionSheetChanged,
  } = input

  const [editorValue, setEditorValue] = useState('')
  const [editorSelectionBehavior, setEditorSelectionBehavior] = useState<EditSelectionBehavior>('select-all')
  const [editorTargetSelection, setEditorTargetSelection] = useState<EditTargetSelection | null>(null)
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
  const optimisticCellSnapshotsRef = useRef<Map<string, CellSnapshot>>(new Map())
  const editSessionRef = useRef(0)
  const editorBaseHydrationPendingRef = useRef(false)
  const pendingEditCommitSessionRef = useRef<number | null>(null)
  const pendingEditCommitMovementAppliedRef = useRef(false)
  const { enqueueDeferredEditCommit, flushPendingEditCommit, isEditCommitPending } = useDeferredEditCommits()

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
    if (activeLocalSelection && workerSelectionAcknowledgesSnapshot(activeLocalSelection, selection)) {
      pendingLocalSelectionRef.current = null
    }
    if (activeExternalSelection && workerSelectionAcknowledgesSnapshot(activeExternalSelection, selection)) {
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

  const hasLiveSelectedCellSnapshot = useCallback(
    (nextSelection = selectionRef.current) => {
      const active = workerHandleRef.current
      if (!active) {
        return selectedCell.sheetName === nextSelection.sheetName && selectedCell.address === nextSelection.address
      }
      return active.viewportStore.hasCellSnapshot?.(nextSelection.sheetName, nextSelection.address) ?? true
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
      optimisticCellSnapshotsRef.current.delete(key)
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])
  const supersedeOptimisticCellSeedsForRange = useCallback((range: CellRangeRef): (() => void) | null => {
    const rollback = supersedeOptimisticSeedsInRange(
      range,
      optimisticCellSeedsRef.current,
      optimisticCellResolvedValuesRef.current,
      optimisticCellSnapshotsRef.current,
    )
    if (!rollback) return null
    bumpOptimisticSeedRevision((revision) => revision + 1)
    return () => {
      rollback()
      bumpOptimisticSeedRevision((revision) => revision + 1)
    }
  }, [])
  const supersedeOptimisticCellSeedsForSheet = useCallback((sheetName: string): (() => void) | null => {
    const prefix = `${sheetName}:`
    const removedSeeds: Array<readonly [string, string, string | undefined, CellSnapshot | undefined]> = []

    for (const [key, seed] of optimisticCellSeedsRef.current) {
      if (key.startsWith(prefix)) {
        removedSeeds.push([key, seed, optimisticCellResolvedValuesRef.current.get(key), optimisticCellSnapshotsRef.current.get(key)])
      }
    }

    if (removedSeeds.length === 0) {
      return null
    }

    for (const [key] of removedSeeds) {
      optimisticCellSeedsRef.current.delete(key)
      optimisticCellResolvedValuesRef.current.delete(key)
      optimisticCellSnapshotsRef.current.delete(key)
    }

    bumpOptimisticSeedRevision((revision) => revision + 1)
    return () => {
      for (const [key, seed, resolvedValue, snapshot] of removedSeeds) {
        if (!optimisticCellSeedsRef.current.has(key)) {
          optimisticCellSeedsRef.current.set(key, seed)
        }
        if (resolvedValue !== undefined && !optimisticCellResolvedValuesRef.current.has(key)) {
          optimisticCellResolvedValuesRef.current.set(key, resolvedValue)
        }
        if (snapshot !== undefined && !optimisticCellSnapshotsRef.current.has(key)) {
          optimisticCellSnapshotsRef.current.set(key, snapshot)
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
      const targetKey = optimisticCellKey(targetSelection.sheetName, targetSelection.address)
      const hadPreviousOptimisticSnapshot = optimisticCellSnapshotsRef.current.has(targetKey)
      const previousOptimisticSnapshot = optimisticCellSnapshotsRef.current.get(targetKey)
      const readVisibleCell = (sheetName: string, address: string): CellSnapshot => {
        const optimisticSnapshot = optimisticCellSnapshotsRef.current.get(optimisticCellKey(sheetName, address))
        if (optimisticSnapshot) {
          return optimisticSnapshot
        }
        if (viewportStore) {
          return viewportStore.getCell(sheetName, address)
        }
        if (selectedCell.sheetName === sheetName && selectedCell.address === address) {
          return selectedCell
        }
        return emptyCellSnapshot(sheetName, address)
      }
      const previous = readVisibleCell(targetSelection.sheetName, targetSelection.address)
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
            getCell: readVisibleCell,
          }),
      })
      optimisticCellSnapshotsRef.current.set(targetKey, optimistic)
      viewportStore?.setCellSnapshot(optimistic, { localDirtyMask: LOCAL_CELL_CONTENT_DIRTY_MASK })
      return {
        editorSeed: toEditorValue(optimistic),
        resolvedValue: toResolvedValue(optimistic),
        rollback: (snapshot = previous) => {
          if (hadPreviousOptimisticSnapshot && previousOptimisticSnapshot) {
            optimisticCellSnapshotsRef.current.set(targetKey, previousOptimisticSnapshot)
          } else {
            optimisticCellSnapshotsRef.current.delete(targetKey)
          }
          viewportStore?.setCellSnapshot(createSupersedingCellSnapshot(snapshot, optimistic.version + 1), {
            localDirtyMask: LOCAL_CELL_CONTENT_DIRTY_MASK,
          })
        },
      }
    },
    [selectedCell, workerHandleRef],
  )

  const refreshMountedEditorTargetSnapshot = useCallback(
    (targetSelection: EditTargetSelection, snapshot: CellSnapshot) => {
      workerHandleRef.current?.viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, snapshot.version + 1), {
        force: true,
        forceOptimistic: true,
        localDirtyMask: LOCAL_CELL_CONTENT_DIRTY_MASK,
      })
      const key = optimisticCellKey(targetSelection.sheetName, targetSelection.address)
      optimisticCellSnapshotsRef.current.delete(key)
      optimisticCellSeedsRef.current.delete(key)
      optimisticCellResolvedValuesRef.current.delete(key)
      bumpOptimisticSeedRevision((revision) => revision + 1)
    },
    [workerHandleRef],
  )

  const reconcileOptimisticCellSeeds = useCallback(
    (visible?: { readonly key: string; readonly cell: CellSnapshot }) => {
      const active = workerHandleRef.current
      let removed = false

      for (const [key, seed] of optimisticCellSeedsRef.current) {
        const target = optimisticCellTargetFromKey(key)
        if (!target) {
          continue
        }
        const liveCell =
          visible?.key === key
            ? visible.cell
            : active
              ? active.viewportStore.getCell(target.sheetName, target.address)
              : selectedCell.sheetName === target.sheetName && selectedCell.address === target.address
                ? selectedCell
                : null

        if (!liveCell || isOptimisticSnapshot(liveCell) || seed !== toEditorValue(liveCell)) {
          continue
        }

        optimisticCellSeedsRef.current.delete(key)
        optimisticCellResolvedValuesRef.current.delete(key)
        optimisticCellSnapshotsRef.current.delete(key)
        removed = true
      }

      if (removed) {
        bumpOptimisticSeedRevision((revision) => revision + 1)
      }
    },
    [selectedCell, workerHandleRef],
  )

  const cloneLiveSelectedCell = useCallback(
    (nextSelection = selectionRef.current) => structuredClone(getLiveSelectedCell(nextSelection)),
    [getLiveSelectedCell],
  )

  const resolveEditorBaseSnapshot = useCallback(
    (targetSelection: WorkerRuntimeSelection, liveSnapshot = cloneLiveSelectedCell(targetSelection)): CellSnapshot => {
      const baseSnapshot = editorBaseSnapshotRef.current
      const baseMatchesTarget = baseSnapshot.sheetName === targetSelection.sheetName && baseSnapshot.address === targetSelection.address
      const hasHydratedTarget = hasLiveSelectedCellSnapshot(targetSelection)
      if (!baseMatchesTarget) {
        editorBaseSnapshotRef.current = liveSnapshot
        editorBaseHydrationPendingRef.current = !hasHydratedTarget
        return liveSnapshot
      }
      if (editorBaseHydrationPendingRef.current && hasHydratedTarget) {
        editorBaseSnapshotRef.current = liveSnapshot
        editorBaseHydrationPendingRef.current = false
        return liveSnapshot
      }
      if (hasHydratedTarget) {
        editorBaseHydrationPendingRef.current = false
      }
      return baseSnapshot
    },
    [cloneLiveSelectedCell, hasLiveSelectedCellSnapshot],
  )

  const resetEditorConflictTracking = useCallback(
    (nextSelection = selectionRef.current) => {
      editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextSelection)
      editorBaseHydrationPendingRef.current = !hasLiveSelectedCellSnapshot(nextSelection)
      setEditorConflict(null)
    },
    [cloneLiveSelectedCell, hasLiveSelectedCellSnapshot],
  )

  const completeSelectionNavigation = useCallback(
    (targetSelection: WorkerRuntimeSelection, movement?: EditMovement) => {
      if (!movement) {
        selectionRef.current = targetSelection
        editorTargetRef.current = targetSelection
        const nextSelectionSnapshot = createSingleCellSelectionSnapshot(targetSelection)
        if (!selectionSnapshotsEqual(selectionSnapshotRef.current, nextSelectionSnapshot)) {
          selectionSnapshotRef.current = nextSelectionSnapshot
          selectionRangeRef.current = selectionSnapshotToRangeRef(nextSelectionSnapshot)
          setSelectionSnapshot(nextSelectionSnapshot)
        }
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
      setEditorTargetSelection(null)
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
      editorBaseHydrationPendingRef.current = !hasLiveSelectedCellSnapshot(nextTarget)
      setEditorConflict(null)
      editorValueRef.current = nextEditorValue
      setEditorValue(nextEditorValue)
      setEditorSelectionBehavior(selectionBehavior)
      editorTargetRef.current = nextTarget
      setEditorTargetSelection(nextTarget)
      editingModeRef.current = mode
      setEditingMode(mode)
    },
    [cloneLiveSelectedCell, getLiveSelectedCell, hasLiveSelectedCellSnapshot, onSelectionSheetChanged, sendSelectionChanged, writesAllowed],
  )

  const applyParsedInput = useCallback(
    async (sheetName: string, address: string, parsed: ParsedEditorInput) => {
      if (parsed.kind === 'formula') {
        await invokeEditCommitMutation('setCellFormula', sheetName, address, parsed.formula)
        perfSession.markFirstLocalEditApplied?.()
        return
      }
      if (parsed.kind === 'clear') {
        await invokeEditCommitMutation('clearCell', sheetName, address)
        perfSession.markFirstLocalEditApplied?.()
        return
      }
      await invokeEditCommitMutation('setCellValue', sheetName, address, parsed.value)
      perfSession.markFirstLocalEditApplied?.()
    },
    [invokeEditCommitMutation, perfSession],
  )

  const commitEditor = useCallback(
    (movement?: EditMovement, valueOverride?: string, targetSelectionOverride?: EditTargetSelection): boolean => {
      if (!writesAllowed) {
        return false
      }
      if (editingModeRef.current === 'idle' && valueOverride === undefined && targetSelectionOverride === undefined) {
        if (movement && !pendingEditCommitMovementAppliedRef.current) {
          pendingEditCommitMovementAppliedRef.current = true
          completeSelectionNavigation(selectionRef.current, movement)
        }
        return true
      }
      const mountedCellEditorValue = editingModeRef.current === 'cell' ? readMountedCellEditorValue() : null
      const targetSelection =
        targetSelectionOverride ?? (editingModeRef.current === 'idle' ? selectionRef.current : editorTargetRef.current)
      const nextValue =
        valueOverride ??
        (editingModeRef.current === 'idle'
          ? toEditorValue(getLiveSelectedCell(targetSelection))
          : (mountedCellEditorValue ?? editorValueRef.current))
      const commitSessionId = editSessionRef.current
      if (pendingEditCommitSessionRef.current === commitSessionId) {
        if (movement && !pendingEditCommitMovementAppliedRef.current) {
          pendingEditCommitMovementAppliedRef.current = true
          finishEditingWithAuthoritative(targetSelection, movement)
        }
        return true
      }
      pendingEditCommitSessionRef.current = commitSessionId
      pendingEditCommitMovementAppliedRef.current = false
      const parsed = parseEditorInput(nextValue)
      const isFreshValueOverride = valueOverride !== undefined && editingModeRef.current === 'idle'
      const liveSnapshot = cloneLiveSelectedCell(targetSelection)
      const baseSnapshot = resolveEditorBaseSnapshot(targetSelection, liveSnapshot)
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
          return true
        }
        setEditorConflict({
          sheetName: targetSelection.sheetName,
          address: targetSelection.address,
          phase: 'compare',
          baseSnapshot: targetBaseSnapshot,
          authoritativeSnapshot: liveSnapshot,
        })
        return false
      }

      if (draftMatchesLiveSnapshot) {
        if (editingModeRef.current === 'cell' && mountedCellEditorValue !== null) {
          refreshMountedEditorTargetSnapshot(targetSelection, liveSnapshot)
        }
        pendingEditCommitSessionRef.current = null
        pendingEditCommitMovementAppliedRef.current = Boolean(movement)
        finishEditingWithAuthoritative(targetSelection, movement)
        return true
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
      finishEditingAtSelection(nextSelection)
      enqueueDeferredEditCommit(async () => {
        try {
          await applyParsedInput(targetSelection.sheetName, targetSelection.address, parsed)
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
          setEditorTargetSelection(null)
          setEditingMode('idle')
          reportRuntimeError(error)
        } finally {
          if (pendingEditCommitSessionRef.current === commitSessionId) {
            pendingEditCommitSessionRef.current = null
          }
        }
      })
      return true
    },
    [
      applyOptimisticParsedInput,
      applyParsedInput,
      clearOptimisticCellSeed,
      cloneLiveSelectedCell,
      completeSelectionNavigation,
      enqueueDeferredEditCommit,
      finishEditingAtSelection,
      finishEditingWithAuthoritative,
      getLiveSelectedCell,
      refreshMountedEditorTargetSnapshot,
      reportRuntimeError,
      resolveEditorBaseSnapshot,
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
    setEditorTargetSelection(null)
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
      const previousSnapshot = selectionSnapshotRef.current
      const wasEditing = editingModeRef.current !== 'idle'
      if (wasEditing) {
        commitEditor(undefined, undefined, editorTargetRef.current)
        if (editingModeRef.current !== 'idle') {
          return
        }
      }
      if (selectionSnapshotsEqual(previousSnapshot, nextSelectionSnapshot)) {
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
        setEditorTargetSelection(nextTarget)
        editorBaseSnapshotRef.current = cloneLiveSelectedCell(nextTarget)
        editorBaseHydrationPendingRef.current = !hasLiveSelectedCellSnapshot(nextTarget)
        setEditorConflict(null)
      }
      editorValueRef.current = next
      setEditorValue(next)
      if (editingModeRef.current === 'idle') {
        editingModeRef.current = 'cell'
        setEditingMode('cell')
      }
    },
    [cloneLiveSelectedCell, hasLiveSelectedCellSnapshot],
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
    reconcileOptimisticCellSeeds({ key: visibleCellKey, cell: liveVisibleCell })
  }, [liveVisibleCell, reconcileOptimisticCellSeeds, visibleCellKey])

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
    resolveEditorBaseSnapshot,
    completeEditNavigation: completeSelectionNavigation,
    finishEditingWithAuthoritative,
    resetEditorConflictTracking,
    applyParsedInput,
    reportRuntimeError,
    setEditorSelectionBehavior,
    setEditorTargetSelection,
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
    editorTargetSelection,
    fillSelectionRange,
    flushPendingEditCommit,
    getCellEditorSeed,
    acknowledgeExternalSelectionSync,
    handleEditorChange,
    handleSelectionChange,
    isEditing,
    isEditingCell,
    isEditCommitPending,
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
