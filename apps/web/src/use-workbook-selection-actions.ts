import { useCallback, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import type { CommitOp } from '@bilig/core'
import { formatAddress, parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { EditSelectionBehavior } from '@bilig/grid'
import { ValueTag, type CellRangeRef, type CellSnapshot } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import { parseEditorInput, parsedEditorInputFromSnapshot, type EditingMode, type ParsedEditorInput } from './worker-workbook-app-model.js'
import { createOptimisticCellSnapshot, createSupersedingCellSnapshot, evaluateOptimisticFormula } from './workbook-optimistic-cell.js'

type RangeMutationMethod = 'fillRange' | 'copyRange' | 'moveRange'

interface OptimisticViewportStore {
  getCell(sheetName: string, address: string): CellSnapshot
  setCellSnapshot(snapshot: CellSnapshot): void
}

export function buildPasteCommitOps(sheetName: string, startAddr: string, values: readonly (readonly string[])[]): CommitOp[] {
  const start = parseCellAddress(startAddr, sheetName)
  const ops: CommitOp[] = []
  values.forEach((rowValues, rowOffset) => {
    rowValues.forEach((cellValue, colOffset) => {
      const address = formatAddress(start.row + rowOffset, start.col + colOffset)
      const parsed = parseEditorInput(cellValue)
      if (parsed.kind === 'formula') {
        ops.push({
          kind: 'upsertCell',
          sheetName,
          addr: address,
          formula: parsed.formula,
        })
        return
      }
      if (parsed.kind === 'clear') {
        ops.push({ kind: 'deleteCell', sheetName, addr: address })
        return
      }
      ops.push({
        kind: 'upsertCell',
        sheetName,
        addr: address,
        value: parsed.value,
      })
    })
  })
  return ops
}

export function createSheetScopedRangePair(
  sheetName: string,
  sourceStartAddr: string,
  sourceEndAddr: string,
  targetStartAddr: string,
  targetEndAddr: string,
): { source: CellRangeRef; target: CellRangeRef } {
  return {
    source: {
      sheetName,
      startAddress: sourceStartAddr,
      endAddress: sourceEndAddr,
    },
    target: {
      sheetName,
      startAddress: targetStartAddr,
      endAddress: targetEndAddr,
    },
  }
}

function applyOptimisticMoveRange(
  viewportStore: OptimisticViewportStore | null,
  source: CellRangeRef,
  target: CellRangeRef,
): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const sourceBounds = normalizeCellRange(source)
  const targetBounds = normalizeCellRange(target)
  const height = sourceBounds.endRow - sourceBounds.startRow + 1
  const width = sourceBounds.endCol - sourceBounds.startCol + 1
  if (height !== targetBounds.endRow - targetBounds.startRow + 1 || width !== targetBounds.endCol - targetBounds.startCol + 1) {
    return null
  }

  const previousSnapshots: CellSnapshot[] = []
  const nextSourceSnapshots: CellSnapshot[] = []
  const nextTargetSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
      const targetAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
      const sourceSnapshot = viewportStore.getCell(source.sheetName, sourceAddress)
      const targetSnapshot = viewportStore.getCell(target.sheetName, targetAddress)
      previousSnapshots.push(sourceSnapshot, targetSnapshot)
      const nextSourceSnapshot = createEmptyOptimisticSnapshot(source.sheetName, sourceAddress, sourceSnapshot.version + 1)
      const nextTargetSnapshot = {
        ...sourceSnapshot,
        sheetName: target.sheetName,
        address: targetAddress,
        version: Math.max(sourceSnapshot.version, targetSnapshot.version) + 1,
      }
      rollbackVersion = Math.max(rollbackVersion, nextSourceSnapshot.version, nextTargetSnapshot.version)
      nextSourceSnapshots.push(nextSourceSnapshot)
      nextTargetSnapshots.push(nextTargetSnapshot)
    }
  }

  nextSourceSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))
  nextTargetSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}

export function applyOptimisticCopyRange(
  viewportStore: OptimisticViewportStore | null,
  source: CellRangeRef,
  target: CellRangeRef,
): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const sourceBounds = normalizeCellRange(source)
  const targetBounds = normalizeCellRange(target)
  const height = sourceBounds.endRow - sourceBounds.startRow + 1
  const width = sourceBounds.endCol - sourceBounds.startCol + 1
  if (height !== targetBounds.endRow - targetBounds.startRow + 1 || width !== targetBounds.endCol - targetBounds.startCol + 1) {
    return null
  }

  const previousSnapshots: CellSnapshot[] = []
  const stagedSnapshots = new Map<string, CellSnapshot>()
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
      const targetAddress = formatAddress(targetBounds.startRow + rowOffset, targetBounds.startCol + colOffset)
      const sourceSnapshot = viewportStore.getCell(source.sheetName, sourceAddress)
      const targetSnapshot = viewportStore.getCell(target.sheetName, targetAddress)
      const parsed = parsedInputForCopiedSnapshot(sourceSnapshot, sourceAddress, target.sheetName, targetAddress)
      const next = createOptimisticCellSnapshot({
        sheetName: target.sheetName,
        address: targetAddress,
        current: targetSnapshot,
        parsed,
        evaluateFormula: (formula) =>
          evaluateOptimisticFormula({
            sheetName: target.sheetName,
            address: targetAddress,
            formula,
            getCell: (sheetName, address) =>
              stagedSnapshots.get(optimisticSnapshotKey(sheetName, address)) ?? viewportStore.getCell(sheetName, address),
          }),
      })
      previousSnapshots.push(targetSnapshot)
      stagedSnapshots.set(optimisticSnapshotKey(target.sheetName, targetAddress), next)
      nextSnapshots.push(next)
      rollbackVersion = Math.max(rollbackVersion, next.version)
    }
  }

  nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}

export function applyOptimisticFillRange(
  viewportStore: OptimisticViewportStore | null,
  source: CellRangeRef,
  target: CellRangeRef,
): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const sourceBounds = normalizeCellRange(source)
  const targetBounds = normalizeCellRange(target)
  const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
  const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
  if (sourceHeight <= 0 || sourceWidth <= 0) {
    return null
  }

  const sourceCells: CellSnapshot[][] = []
  for (let rowOffset = 0; rowOffset < sourceHeight; rowOffset += 1) {
    const rowCells: CellSnapshot[] = []
    for (let colOffset = 0; colOffset < sourceWidth; colOffset += 1) {
      const sourceAddress = formatAddress(sourceBounds.startRow + rowOffset, sourceBounds.startCol + colOffset)
      rowCells.push(viewportStore.getCell(source.sheetName, sourceAddress))
    }
    sourceCells.push(rowCells)
  }

  const previousSnapshots: CellSnapshot[] = []
  const stagedSnapshots = new Map<string, CellSnapshot>()
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  for (let row = targetBounds.startRow; row <= targetBounds.endRow; row += 1) {
    for (let col = targetBounds.startCol; col <= targetBounds.endCol; col += 1) {
      const sourceRowOffset = (row - targetBounds.startRow) % sourceHeight
      const sourceColOffset = (col - targetBounds.startCol) % sourceWidth
      const sourceAddress = formatAddress(sourceBounds.startRow + sourceRowOffset, sourceBounds.startCol + sourceColOffset)
      const targetAddress = formatAddress(row, col)
      const sourceSnapshot = sourceCells[sourceRowOffset]?.[sourceColOffset]
      if (!sourceSnapshot) {
        continue
      }
      const targetSnapshot = viewportStore.getCell(target.sheetName, targetAddress)
      const parsed = parsedInputForCopiedSnapshot(sourceSnapshot, sourceAddress, target.sheetName, targetAddress)
      const next = createOptimisticCellSnapshot({
        sheetName: target.sheetName,
        address: targetAddress,
        current: targetSnapshot,
        parsed,
        evaluateFormula: (formula) =>
          evaluateOptimisticFormula({
            sheetName: target.sheetName,
            address: targetAddress,
            formula,
            getCell: (sheetName, address) =>
              stagedSnapshots.get(optimisticSnapshotKey(sheetName, address)) ?? viewportStore.getCell(sheetName, address),
          }),
      })
      previousSnapshots.push(targetSnapshot)
      stagedSnapshots.set(optimisticSnapshotKey(target.sheetName, targetAddress), next)
      nextSnapshots.push(next)
      rollbackVersion = Math.max(rollbackVersion, next.version)
    }
  }

  if (nextSnapshots.length === 0) {
    return null
  }

  nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}

export function applyOptimisticClearRange(viewportStore: OptimisticViewportStore | null, range: CellRangeRef): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const bounds = normalizeCellRange(range)
  const previousSnapshots: CellSnapshot[] = []
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let col = bounds.startCol; col <= bounds.endCol; col += 1) {
      const address = formatAddress(row, col)
      const previous = viewportStore.getCell(range.sheetName, address)
      const next = createEmptyOptimisticSnapshot(range.sheetName, address, previous.version + 1)
      previousSnapshots.push(previous)
      nextSnapshots.push(next)
      rollbackVersion = Math.max(rollbackVersion, next.version)
    }
  }

  nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}

export function applyOptimisticCommitOps(viewportStore: OptimisticViewportStore | null, ops: readonly CommitOp[]): (() => void) | null {
  if (!viewportStore) {
    return null
  }

  const previousSnapshots: CellSnapshot[] = []
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  for (const op of ops) {
    if ((op.kind !== 'upsertCell' && op.kind !== 'deleteCell') || !op.sheetName || !op.addr) {
      continue
    }
    const opSheetName = op.sheetName
    const opAddress = op.addr
    const previous = viewportStore.getCell(opSheetName, opAddress)
    previousSnapshots.push(previous)
    const parsed = parsedInputFromCommitOp(op)
    const next =
      parsed.kind === 'clear'
        ? createEmptyOptimisticSnapshot(opSheetName, opAddress, previous.version + 1)
        : createOptimisticCellSnapshot({
            sheetName: opSheetName,
            address: opAddress,
            current: previous,
            parsed,
            evaluateFormula: (formula) =>
              evaluateOptimisticFormula({
                sheetName: opSheetName,
                address: opAddress,
                formula,
                getCell: (refSheetName, refAddress) => viewportStore.getCell(refSheetName, refAddress),
              }),
          })
    rollbackVersion = Math.max(rollbackVersion, next.version)
    nextSnapshots.push(next)
  }

  if (nextSnapshots.length === 0) {
    return null
  }

  nextSnapshots.forEach((snapshot) => viewportStore.setCellSnapshot(snapshot))

  return () => {
    previousSnapshots.forEach((snapshot) => {
      rollbackVersion += 1
      viewportStore.setCellSnapshot(createSupersedingCellSnapshot(snapshot, rollbackVersion))
    })
  }
}

function normalizeCellRange(range: CellRangeRef): { startRow: number; endRow: number; startCol: number; endCol: number } {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  }
}

function parsedInputForCopiedSnapshot(
  snapshot: CellSnapshot,
  sourceAddress: string,
  targetSheetName: string,
  targetAddress: string,
): ParsedEditorInput {
  if (typeof snapshot.formula === 'string') {
    return {
      kind: 'formula',
      formula:
        snapshot.sheetName === targetSheetName
          ? translateFormulaReferences(
              snapshot.formula,
              parseCellAddress(targetAddress, targetSheetName).row - parseCellAddress(sourceAddress, snapshot.sheetName).row,
              parseCellAddress(targetAddress, targetSheetName).col - parseCellAddress(sourceAddress, snapshot.sheetName).col,
            )
          : snapshot.formula,
    }
  }
  return parsedEditorInputFromSnapshot(snapshot)
}

function optimisticSnapshotKey(sheetName: string, address: string): string {
  return `${sheetName}:${address}`
}

function parsedInputFromCommitOp(op: CommitOp): ParsedEditorInput {
  if (op.kind === 'deleteCell') {
    return { kind: 'clear' }
  }
  if (typeof op.formula === 'string') {
    return { kind: 'formula', formula: op.formula }
  }
  return { kind: 'value', value: op.value ?? null }
}

function createEmptyOptimisticSnapshot(sheetName: string, address: string, version: number): CellSnapshot {
  return {
    sheetName,
    address,
    value: { tag: ValueTag.Empty },
    flags: 0,
    version,
  }
}

export function useWorkbookSelectionActions(input: {
  writesAllowed: boolean
  selectionRangeRef: MutableRefObject<CellRangeRef>
  selectionRef: MutableRefObject<WorkerRuntimeSelection>
  editorTargetRef: MutableRefObject<WorkerRuntimeSelection>
  editorValueRef: MutableRefObject<string>
  editingModeRef: MutableRefObject<EditingMode>
  invokeMutation: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>
  applyParsedInput: (sheetName: string, address: string, parsed: ParsedEditorInput) => Promise<void>
  viewportStore?: OptimisticViewportStore | null | undefined
  onPasteApplied?: () => void
  resetEditorConflictTracking: (nextSelection?: WorkerRuntimeSelection) => void
  reportRuntimeError: (error: unknown) => void
  setEditorValue: Dispatch<SetStateAction<string>>
  setEditingMode: Dispatch<SetStateAction<EditingMode>>
  setEditorSelectionBehavior: Dispatch<SetStateAction<EditSelectionBehavior>>
}) {
  const {
    applyParsedInput,
    editingModeRef,
    editorTargetRef,
    editorValueRef,
    invokeMutation,
    onPasteApplied,
    reportRuntimeError,
    resetEditorConflictTracking,
    selectionRangeRef,
    selectionRef,
    setEditingMode,
    setEditorSelectionBehavior,
    setEditorValue,
    viewportStore,
    writesAllowed,
  } = input

  const resetEditingState = useCallback(
    (nextEditorValue?: string) => {
      if (nextEditorValue !== undefined) {
        editorValueRef.current = nextEditorValue
        setEditorValue(nextEditorValue)
      }
      setEditorSelectionBehavior('select-all')
      editorTargetRef.current = selectionRef.current
      editingModeRef.current = 'idle'
      setEditingMode('idle')
    },
    [editingModeRef, editorTargetRef, editorValueRef, selectionRef, setEditingMode, setEditorSelectionBehavior, setEditorValue],
  )

  const runRangeMutation = useCallback(
    (method: RangeMutationMethod, sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
      if (!writesAllowed) {
        return
      }
      const { source, target } = createSheetScopedRangePair(
        selectionRef.current.sheetName,
        sourceStartAddr,
        sourceEndAddr,
        targetStartAddr,
        targetEndAddr,
      )
      const rollbackOptimisticRange =
        method === 'moveRange'
          ? applyOptimisticMoveRange(viewportStore ?? null, source, target)
          : method === 'copyRange'
            ? applyOptimisticCopyRange(viewportStore ?? null, source, target)
            : applyOptimisticFillRange(viewportStore ?? null, source, target)
      void (async () => {
        try {
          await invokeMutation(method, source, target)
          resetEditingState()
          resetEditorConflictTracking()
        } catch (error) {
          rollbackOptimisticRange?.()
          reportRuntimeError(error)
        }
      })()
    },
    [invokeMutation, reportRuntimeError, resetEditingState, resetEditorConflictTracking, selectionRef, viewportStore, writesAllowed],
  )

  const clearSelectedRange = useCallback(() => {
    if (!writesAllowed) {
      return
    }
    const range = selectionRangeRef.current
    const rollbackOptimisticClear = applyOptimisticClearRange(viewportStore ?? null, range)
    resetEditingState('')
    resetEditorConflictTracking()
    void (async () => {
      try {
        await invokeMutation('clearRange', range)
      } catch (error) {
        rollbackOptimisticClear?.()
        reportRuntimeError(error)
      }
    })()
  }, [invokeMutation, reportRuntimeError, resetEditingState, resetEditorConflictTracking, selectionRangeRef, viewportStore, writesAllowed])

  const clearSelectedCell = useCallback(() => {
    clearSelectedRange()
  }, [clearSelectedRange])

  const toggleBooleanCell = useCallback(
    (sheetName: string, address: string, nextValue: boolean) => {
      if (!writesAllowed) {
        return
      }
      void (async () => {
        try {
          await applyParsedInput(sheetName, address, { kind: 'value', value: nextValue })
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    [applyParsedInput, reportRuntimeError, writesAllowed],
  )

  const pasteIntoSelection = useCallback(
    (sheetName: string, startAddr: string, values: readonly (readonly string[])[]) => {
      if (!writesAllowed) {
        return
      }
      const ops = buildPasteCommitOps(sheetName, startAddr, values)
      if (ops.length === 0) {
        return
      }
      const rollbackOptimisticPaste = applyOptimisticCommitOps(viewportStore ?? null, ops)
      void (async () => {
        try {
          await invokeMutation('renderCommit', ops)
          onPasteApplied?.()
        } catch (error) {
          rollbackOptimisticPaste?.()
          reportRuntimeError(error)
        }
      })()
      resetEditingState()
      resetEditorConflictTracking()
    },
    [invokeMutation, onPasteApplied, reportRuntimeError, resetEditingState, resetEditorConflictTracking, viewportStore, writesAllowed],
  )

  const fillSelectionRange = useCallback(
    (sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
      runRangeMutation('fillRange', sourceStartAddr, sourceEndAddr, targetStartAddr, targetEndAddr)
    },
    [runRangeMutation],
  )

  const copySelectionRange = useCallback(
    (sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
      runRangeMutation('copyRange', sourceStartAddr, sourceEndAddr, targetStartAddr, targetEndAddr)
    },
    [runRangeMutation],
  )

  const moveSelectionRange = useCallback(
    (sourceStartAddr: string, sourceEndAddr: string, targetStartAddr: string, targetEndAddr: string) => {
      runRangeMutation('moveRange', sourceStartAddr, sourceEndAddr, targetStartAddr, targetEndAddr)
    },
    [runRangeMutation],
  )

  return {
    clearSelectedCell,
    clearSelectedRange,
    copySelectionRange,
    fillSelectionRange,
    moveSelectionRange,
    pasteIntoSelection,
    toggleBooleanCell,
  }
}
