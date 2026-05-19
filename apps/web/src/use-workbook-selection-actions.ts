import { useCallback, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import type { CommitOp } from '@bilig/core'
import { formatAddress, parseCellAddress, translateFormulaReferences } from '@bilig/formula'
import type { EditSelectionBehavior, GridSelectionSnapshot } from '@bilig/grid'
import type { CellRangeRef, CellSnapshot } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'
import { parseEditorInput, parsedEditorInputFromSnapshot, type EditingMode, type ParsedEditorInput } from './worker-workbook-app-model.js'
import { createOptimisticCellSnapshot, createSupersedingCellSnapshot, evaluateOptimisticFormula } from './workbook-optimistic-cell.js'
import { createEmptyOptimisticSnapshot, normalizeCellRange, type OptimisticViewportStore } from './workbook-optimistic-range.js'

export { applyOptimisticClearRange } from './workbook-optimistic-range.js'

type RangeMutationMethod = 'fillRange' | 'copyRange' | 'moveRange'
const MAX_MATERIALIZED_OPTIMISTIC_RANGE_CELLS = 10_000

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

export function applyOptimisticMoveRange(
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

  if (height * width > MAX_MATERIALIZED_OPTIMISTIC_RANGE_CELLS) {
    return applyVisibleOptimisticMoveRange(viewportStore, source, target, sourceBounds, targetBounds)
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
        flags: sourceSnapshot.flags | OPTIMISTIC_CELL_SNAPSHOT_FLAG,
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

  if (height * width > MAX_MATERIALIZED_OPTIMISTIC_RANGE_CELLS) {
    return applyVisibleOptimisticCopyRange(viewportStore, source, target, sourceBounds, targetBounds)
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
      applyCopiedSnapshotPresentation(next, sourceSnapshot)
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
  const targetCellCount = (targetBounds.endRow - targetBounds.startRow + 1) * (targetBounds.endCol - targetBounds.startCol + 1)
  if (targetCellCount > MAX_MATERIALIZED_OPTIMISTIC_RANGE_CELLS) {
    return (
      applyOptimisticFillRangeOverlay(viewportStore, source, target, sourceBounds, targetBounds) ??
      applyVisibleOptimisticFillRange(viewportStore, source, target, sourceBounds, targetBounds)
    )
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
      applyCopiedSnapshotPresentation(next, sourceSnapshot)
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

function applyVisibleOptimisticMoveRange(
  viewportStore: OptimisticViewportStore,
  source: CellRangeRef,
  target: CellRangeRef,
  sourceBounds: ReturnType<typeof normalizeCellRange>,
  targetBounds: ReturnType<typeof normalizeCellRange>,
): (() => void) | null {
  if (!viewportStore.forEachCachedOrVisibleCellSnapshotInRange) {
    return null
  }

  const previousSnapshots: CellSnapshot[] = []
  const nextSourceSnapshots: CellSnapshot[] = []
  const nextTargetSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0

  viewportStore.forEachCachedOrVisibleCellSnapshotInRange(source, (sourceSnapshot) => {
    const next = createEmptyOptimisticSnapshot(sourceSnapshot.sheetName, sourceSnapshot.address, sourceSnapshot.version + 1)
    previousSnapshots.push(sourceSnapshot)
    nextSourceSnapshots.push(next)
    rollbackVersion = Math.max(rollbackVersion, next.version)
  })
  viewportStore.forEachCachedOrVisibleCellSnapshotInRange(target, (targetSnapshot) => {
    const targetPosition = parseCellAddress(targetSnapshot.address, targetSnapshot.sheetName)
    const sourceAddress = formatAddress(
      sourceBounds.startRow + (targetPosition.row - targetBounds.startRow),
      sourceBounds.startCol + (targetPosition.col - targetBounds.startCol),
    )
    const sourceSnapshot = getVisibleOptimisticSourceSnapshot(viewportStore, source.sheetName, sourceAddress)
    if (!sourceSnapshot) {
      return
    }
    const next = {
      ...sourceSnapshot,
      sheetName: target.sheetName,
      address: targetSnapshot.address,
      flags: sourceSnapshot.flags | OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: Math.max(sourceSnapshot.version, targetSnapshot.version) + 1,
    }
    previousSnapshots.push(targetSnapshot)
    nextTargetSnapshots.push(next)
    rollbackVersion = Math.max(rollbackVersion, next.version)
  })

  if (nextSourceSnapshots.length + nextTargetSnapshots.length === 0) {
    return null
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

function applyVisibleOptimisticCopyRange(
  viewportStore: OptimisticViewportStore,
  source: CellRangeRef,
  target: CellRangeRef,
  sourceBounds: ReturnType<typeof normalizeCellRange>,
  targetBounds: ReturnType<typeof normalizeCellRange>,
): (() => void) | null {
  if (!viewportStore.forEachCachedOrVisibleCellSnapshotInRange) {
    return null
  }

  const stagedSnapshots = new Map<string, CellSnapshot>()
  return applyVisibleOptimisticTargetRange(viewportStore, target, (targetSnapshot) => {
    const targetPosition = parseCellAddress(targetSnapshot.address, targetSnapshot.sheetName)
    const sourceAddress = formatAddress(
      sourceBounds.startRow + (targetPosition.row - targetBounds.startRow),
      sourceBounds.startCol + (targetPosition.col - targetBounds.startCol),
    )
    const sourceSnapshot = getVisibleOptimisticSourceSnapshot(viewportStore, source.sheetName, sourceAddress)
    if (!sourceSnapshot) {
      return null
    }
    const next = createCopiedOptimisticSnapshot(viewportStore, sourceSnapshot, sourceAddress, targetSnapshot, stagedSnapshots)
    stagedSnapshots.set(optimisticSnapshotKey(targetSnapshot.sheetName, targetSnapshot.address), next)
    return next
  })
}

function applyVisibleOptimisticFillRange(
  viewportStore: OptimisticViewportStore,
  source: CellRangeRef,
  target: CellRangeRef,
  sourceBounds: ReturnType<typeof normalizeCellRange>,
  targetBounds: ReturnType<typeof normalizeCellRange>,
): (() => void) | null {
  if (!viewportStore.forEachCachedOrVisibleCellSnapshotInRange) {
    return null
  }

  const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
  const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
  const stagedSnapshots = new Map<string, CellSnapshot>()
  return applyVisibleOptimisticTargetRange(viewportStore, target, (targetSnapshot) => {
    const targetPosition = parseCellAddress(targetSnapshot.address, targetSnapshot.sheetName)
    const sourceRowOffset = (targetPosition.row - targetBounds.startRow) % sourceHeight
    const sourceColOffset = (targetPosition.col - targetBounds.startCol) % sourceWidth
    const sourceAddress = formatAddress(sourceBounds.startRow + sourceRowOffset, sourceBounds.startCol + sourceColOffset)
    const sourceSnapshot = getVisibleOptimisticSourceSnapshot(viewportStore, source.sheetName, sourceAddress)
    if (!sourceSnapshot) {
      return null
    }
    const next = createCopiedOptimisticSnapshot(viewportStore, sourceSnapshot, sourceAddress, targetSnapshot, stagedSnapshots)
    stagedSnapshots.set(optimisticSnapshotKey(targetSnapshot.sheetName, targetSnapshot.address), next)
    return next
  })
}

function applyOptimisticFillRangeOverlay(
  viewportStore: OptimisticViewportStore,
  source: CellRangeRef,
  target: CellRangeRef,
  sourceBounds: ReturnType<typeof normalizeCellRange>,
  targetBounds: ReturnType<typeof normalizeCellRange>,
): (() => void) | null {
  if (!viewportStore.beginOptimisticRangeOverlay) {
    return null
  }
  const sourceSnapshots = collectCompleteVisibleSourceSnapshots(viewportStore, source, sourceBounds)
  if (!sourceSnapshots) {
    return null
  }
  const sourceHeight = sourceBounds.endRow - sourceBounds.startRow + 1
  const sourceWidth = sourceBounds.endCol - sourceBounds.startCol + 1
  return viewportStore.beginOptimisticRangeOverlay(target, (targetSnapshot) => {
    const targetPosition = parseCellAddress(targetSnapshot.address, targetSnapshot.sheetName)
    const sourceRowOffset = (targetPosition.row - targetBounds.startRow) % sourceHeight
    const sourceColOffset = (targetPosition.col - targetBounds.startCol) % sourceWidth
    const sourceAddress = formatAddress(sourceBounds.startRow + sourceRowOffset, sourceBounds.startCol + sourceColOffset)
    const sourceSnapshot = sourceSnapshots.get(optimisticSnapshotKey(source.sheetName, sourceAddress))
    if (!sourceSnapshot) {
      return targetSnapshot
    }
    return createCopiedOptimisticSnapshot(viewportStore, sourceSnapshot, sourceAddress, targetSnapshot, new Map())
  })
}

function collectCompleteVisibleSourceSnapshots(
  viewportStore: OptimisticViewportStore,
  source: CellRangeRef,
  sourceBounds: ReturnType<typeof normalizeCellRange>,
): Map<string, CellSnapshot> | null {
  const sourceCellCount = (sourceBounds.endRow - sourceBounds.startRow + 1) * (sourceBounds.endCol - sourceBounds.startCol + 1)
  if (sourceCellCount > MAX_MATERIALIZED_OPTIMISTIC_RANGE_CELLS) {
    return null
  }
  const snapshots = new Map<string, CellSnapshot>()
  if (viewportStore.forEachCachedOrVisibleCellSnapshotInRange) {
    viewportStore.forEachCachedOrVisibleCellSnapshotInRange(source, (snapshot) => {
      snapshots.set(optimisticSnapshotKey(snapshot.sheetName, snapshot.address), snapshot)
    })
  }
  for (let row = sourceBounds.startRow; row <= sourceBounds.endRow; row += 1) {
    for (let col = sourceBounds.startCol; col <= sourceBounds.endCol; col += 1) {
      const address = formatAddress(row, col)
      const key = optimisticSnapshotKey(source.sheetName, address)
      if (!snapshots.has(key)) {
        const snapshot = viewportStore.peekCell?.(source.sheetName, address)
        if (!snapshot) {
          return null
        }
        snapshots.set(key, snapshot)
      }
    }
  }
  return snapshots
}

function applyVisibleOptimisticTargetRange(
  viewportStore: OptimisticViewportStore,
  target: CellRangeRef,
  createNextSnapshot: (targetSnapshot: CellSnapshot) => CellSnapshot | null,
): (() => void) | null {
  const previousSnapshots: CellSnapshot[] = []
  const nextSnapshots: CellSnapshot[] = []
  let rollbackVersion = 0
  viewportStore.forEachCachedOrVisibleCellSnapshotInRange?.(target, (targetSnapshot) => {
    const next = createNextSnapshot(targetSnapshot)
    if (!next) {
      return
    }
    previousSnapshots.push(targetSnapshot)
    nextSnapshots.push(next)
    rollbackVersion = Math.max(rollbackVersion, next.version)
  })

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

function getVisibleOptimisticSourceSnapshot(
  viewportStore: OptimisticViewportStore,
  sheetName: string,
  address: string,
): CellSnapshot | null {
  if (!viewportStore.peekCell) {
    return viewportStore.getCell(sheetName, address)
  }
  return viewportStore.peekCell(sheetName, address) ?? null
}

function createCopiedOptimisticSnapshot(
  viewportStore: OptimisticViewportStore,
  sourceSnapshot: CellSnapshot,
  sourceAddress: string,
  targetSnapshot: CellSnapshot,
  stagedSnapshots: ReadonlyMap<string, CellSnapshot>,
): CellSnapshot {
  const parsed = parsedInputForCopiedSnapshot(sourceSnapshot, sourceAddress, targetSnapshot.sheetName, targetSnapshot.address)
  const next = createOptimisticCellSnapshot({
    sheetName: targetSnapshot.sheetName,
    address: targetSnapshot.address,
    current: targetSnapshot,
    parsed,
    evaluateFormula: (formula) =>
      evaluateOptimisticFormula({
        sheetName: targetSnapshot.sheetName,
        address: targetSnapshot.address,
        formula,
        getCell: (sheetName, address) =>
          stagedSnapshots.get(optimisticSnapshotKey(sheetName, address)) ?? viewportStore.getCell(sheetName, address),
      }),
  })
  applyCopiedSnapshotPresentation(next, sourceSnapshot)
  return next
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

function applyCopiedSnapshotPresentation(target: CellSnapshot, source: CellSnapshot): void {
  if (source.format === undefined) {
    delete target.format
  } else {
    target.format = source.format
  }
  if (source.numberFormatId === undefined) {
    delete target.numberFormatId
  } else {
    target.numberFormatId = source.numberFormatId
  }
  if (source.styleId === undefined) {
    delete target.styleId
  } else {
    target.styleId = source.styleId
  }
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

function gridSelectionSnapshotToRangeRef(selection: GridSelectionSnapshot): CellRangeRef {
  return {
    sheetName: selection.sheetName,
    startAddress: selection.range.startAddress,
    endAddress: selection.range.endAddress,
  }
}

function createPasteTargetRange(sheetName: string, startAddr: string, values: readonly (readonly string[])[]): CellRangeRef | null {
  const rowCount = values.length
  const colCount = values.reduce((max, row) => Math.max(max, row.length), 0)
  if (rowCount === 0 || colCount === 0) {
    return null
  }
  const start = parseCellAddress(startAddr, sheetName)
  return {
    sheetName,
    startAddress: startAddr,
    endAddress: formatAddress(start.row + rowCount - 1, start.col + colCount - 1),
  }
}

function combineRollbacks(...rollbacks: Array<(() => void) | null | undefined>): (() => void) | null {
  const activeRollbacks = rollbacks.filter((rollback): rollback is () => void => typeof rollback === 'function')
  if (activeRollbacks.length === 0) {
    return null
  }
  return () => {
    for (let index = activeRollbacks.length - 1; index >= 0; index -= 1) {
      activeRollbacks[index]?.()
    }
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
  supersedeOptimisticCellSeedsForRange?: ((range: CellRangeRef) => (() => void) | null) | undefined
  replaceOptimisticCellSeed?: ((sheetName: string, address: string, seed: string) => (() => void) | null) | undefined
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
    replaceOptimisticCellSeed,
    resetEditorConflictTracking,
    selectionRangeRef,
    selectionRef,
    setEditingMode,
    setEditorSelectionBehavior,
    setEditorValue,
    supersedeOptimisticCellSeedsForRange,
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
      const rollbackOptimisticSeeds = combineRollbacks(
        method === 'moveRange' ? supersedeOptimisticCellSeedsForRange?.(source) : null,
        supersedeOptimisticCellSeedsForRange?.(target),
      )
      void (async () => {
        try {
          await invokeMutation(method, source, target)
          resetEditingState()
          resetEditorConflictTracking()
        } catch (error) {
          rollbackOptimisticSeeds?.()
          reportRuntimeError(error)
        }
      })()
    },
    [
      invokeMutation,
      reportRuntimeError,
      resetEditingState,
      resetEditorConflictTracking,
      selectionRef,
      supersedeOptimisticCellSeedsForRange,
      writesAllowed,
    ],
  )

  const clearSelectedRange = useCallback(
    (targetSelectionSnapshot?: GridSelectionSnapshot) => {
      if (!writesAllowed) {
        return
      }
      const range = targetSelectionSnapshot ? gridSelectionSnapshotToRangeRef(targetSelectionSnapshot) : selectionRangeRef.current
      const activeAddress = targetSelectionSnapshot?.address ?? selectionRef.current.address
      const rollbackOptimisticSeeds = combineRollbacks(
        supersedeOptimisticCellSeedsForRange?.(range),
        replaceOptimisticCellSeed?.(range.sheetName, activeAddress, ''),
      )
      const mutationTask = invokeMutation('clearRange', range)
      resetEditingState('')
      resetEditorConflictTracking()
      void (async () => {
        try {
          await mutationTask
        } catch (error) {
          rollbackOptimisticSeeds?.()
          reportRuntimeError(error)
        }
      })()
    },
    [
      invokeMutation,
      reportRuntimeError,
      resetEditingState,
      resetEditorConflictTracking,
      selectionRangeRef,
      selectionRef,
      replaceOptimisticCellSeed,
      supersedeOptimisticCellSeedsForRange,
      writesAllowed,
    ],
  )

  const clearSelectedCell = useCallback(
    (targetSelectionSnapshot?: GridSelectionSnapshot) => {
      clearSelectedRange(targetSelectionSnapshot)
    },
    [clearSelectedRange],
  )

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
      const targetRange = createPasteTargetRange(sheetName, startAddr, values)
      const rollbackOptimisticSeeds = targetRange ? (supersedeOptimisticCellSeedsForRange?.(targetRange) ?? null) : null
      void (async () => {
        try {
          await invokeMutation('renderCommit', ops)
          onPasteApplied?.()
        } catch (error) {
          rollbackOptimisticSeeds?.()
          reportRuntimeError(error)
        }
      })()
      resetEditingState()
      resetEditorConflictTracking()
    },
    [
      invokeMutation,
      onPasteApplied,
      reportRuntimeError,
      resetEditingState,
      resetEditorConflictTracking,
      supersedeOptimisticCellSeedsForRange,
      writesAllowed,
    ],
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
