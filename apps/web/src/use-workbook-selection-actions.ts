import { useCallback, type Dispatch, type MutableRefObject, type SetStateAction } from 'react'
import type { CommitOp } from '@bilig/core'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { EditSelectionBehavior } from '@bilig/grid'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import type { WorkbookMutationMethod } from './workbook-sync.js'
import { parseEditorInput, type EditingMode, type ParsedEditorInput } from './worker-workbook-app-model.js'

type RangeMutationMethod = 'fillRange' | 'copyRange' | 'moveRange'

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

export function useWorkbookSelectionActions(input: {
  writesAllowed: boolean
  selectionRangeRef: MutableRefObject<CellRangeRef>
  selectionRef: MutableRefObject<WorkerRuntimeSelection>
  editorTargetRef: MutableRefObject<WorkerRuntimeSelection>
  editorValueRef: MutableRefObject<string>
  editingModeRef: MutableRefObject<EditingMode>
  invokeMutation: (method: WorkbookMutationMethod, ...args: unknown[]) => Promise<void>
  applyParsedInput: (sheetName: string, address: string, parsed: ParsedEditorInput) => Promise<void>
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
      void (async () => {
        try {
          await invokeMutation(method, source, target)
          resetEditingState()
          resetEditorConflictTracking()
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    [invokeMutation, reportRuntimeError, resetEditingState, resetEditorConflictTracking, selectionRef, writesAllowed],
  )

  const clearSelectedRange = useCallback(() => {
    if (!writesAllowed) {
      return
    }
    resetEditingState('')
    resetEditorConflictTracking()
    void (async () => {
      try {
        await invokeMutation('clearRange', selectionRangeRef.current)
      } catch (error) {
        reportRuntimeError(error)
      }
    })()
  }, [invokeMutation, reportRuntimeError, resetEditingState, resetEditorConflictTracking, selectionRangeRef, writesAllowed])

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
      void (async () => {
        try {
          await invokeMutation('renderCommit', ops)
          onPasteApplied?.()
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
      resetEditingState()
      resetEditorConflictTracking()
    },
    [invokeMutation, onPasteApplied, reportRuntimeError, resetEditingState, resetEditorConflictTracking, writesAllowed],
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
