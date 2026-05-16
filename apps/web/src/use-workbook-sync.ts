import { useCallback, useEffect, useRef, type MutableRefObject } from 'react'
import { flushSync } from 'react-dom'
import { PRODUCT_COLUMN_WIDTH, PRODUCT_ROW_HEIGHT } from '@bilig/grid'
import type { WorkerHandle, WorkerRuntimeSessionController } from './runtime-session.js'
import {
  buildZeroWorkbookMutation,
  isCellNumberFormatInputValue,
  isCellRangeRef,
  isCellStyleFieldList,
  isCellStylePatchValue,
  isCommitOps,
  isLiteralInput,
  isPendingWorkbookMutation,
  isPendingWorkbookMutationList,
  isWorkbookSheetName,
  isWorkbookStructuralCount,
  isWorkbookStructuralIndex,
  isWorkbookStructuralSize,
  type PendingWorkbookMutation,
  type PendingWorkbookMutationInput,
  type WorkbookMutationMethod,
} from './workbook-sync.js'
import { isPendingWorkbookMutationReadyForSubmission } from './workbook-mutation-journal.js'
import {
  assert,
  canAttemptRemoteSync,
  isMutationErrorResult,
  parsedEditorInputMatchesSnapshot,
  toErrorMessage,
  type ParsedEditorInput,
  type ZeroConnectionState,
} from './worker-workbook-app-model.js'
import { createOptimisticCellSnapshot, createSupersedingCellSnapshot, evaluateOptimisticFormula } from './workbook-optimistic-cell.js'

interface ZeroMutationSource {
  mutate(mutation: unknown): unknown
}

type WorkbookSyncRuntimeController = Pick<WorkerRuntimeSessionController, 'invoke'>
type ViewportStore = WorkerHandle['viewportStore']

const AUTHORITATIVE_REFRESH_PROBE_DELAYS_MS = [400, 1_200, 3_000] as const

type ViewportAxisSizeMutationOptions = {
  deferLocalApplication?: boolean
  flush?: boolean
  deferPersistence?: boolean
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function deferViewportAxisSizeFrame(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(() => resolve())
      return
    }
    setTimeout(resolve, 0)
  })
}

function observeZeroMutationResult(result: unknown): Promise<unknown> | null {
  if (!isRecord(result)) {
    return null
  }
  const observer = result['server'] ?? result['client']
  return observer instanceof Promise ? observer : null
}

function parsedCellInputForMutation(
  mutation: PendingWorkbookMutationInput,
): { readonly sheetName: string; readonly address: string; readonly parsed: ParsedEditorInput } | null {
  if (mutation.method === 'setCellValue') {
    const [sheetName, address, value] = mutation.args
    if (typeof sheetName !== 'string' || typeof address !== 'string' || !isLiteralInput(value)) {
      return null
    }
    return { sheetName, address, parsed: { kind: 'value', value } }
  }
  if (mutation.method === 'setCellFormula') {
    const [sheetName, address, formula] = mutation.args
    if (typeof sheetName !== 'string' || typeof address !== 'string' || typeof formula !== 'string') {
      return null
    }
    return { sheetName, address, parsed: { kind: 'formula', formula } }
  }
  if (mutation.method === 'clearCell') {
    const [sheetName, address] = mutation.args
    if (typeof sheetName !== 'string' || typeof address !== 'string') {
      return null
    }
    return { sheetName, address, parsed: { kind: 'clear' } }
  }
  return null
}

function applyOptimisticCellMutation(
  viewportStore: ViewportStore | null | undefined,
  mutation: PendingWorkbookMutationInput,
): (() => void) | null {
  if (!viewportStore) {
    return null
  }
  const target = parsedCellInputForMutation(mutation)
  if (!target) {
    return null
  }
  const previous = viewportStore.getCell(target.sheetName, target.address)
  if (parsedEditorInputMatchesSnapshot(target.parsed, previous)) {
    return null
  }
  const optimistic = createOptimisticCellSnapshot({
    sheetName: target.sheetName,
    address: target.address,
    current: previous,
    parsed: target.parsed,
    evaluateFormula: (formula) =>
      evaluateOptimisticFormula({
        sheetName: target.sheetName,
        address: target.address,
        formula,
        getCell: (sheetName, address) => viewportStore.getCell(sheetName, address),
      }),
  })
  viewportStore.setCellSnapshot(optimistic)
  return () => {
    viewportStore.setCellSnapshot(createSupersedingCellSnapshot(previous, optimistic.version + 1))
  }
}

export function useWorkbookSync(input: {
  documentId: string
  connectionStateName: ZeroConnectionState['name']
  connectionStateRef: MutableRefObject<ZeroConnectionState['name']>
  runtimeController: WorkbookSyncRuntimeController | null
  workerHandleRef: MutableRefObject<WorkerHandle | null>
  zeroRef: MutableRefObject<ZeroMutationSource>
  reportRuntimeError: (error: unknown) => void
}) {
  const { documentId, connectionStateName, connectionStateRef, runtimeController, workerHandleRef, zeroRef, reportRuntimeError } = input
  const localMutationQueueRef = useRef<Promise<void>>(Promise.resolve())
  const syncQueueRef = useRef<Promise<void>>(Promise.resolve())
  const authoritativeRefreshTimerRefs = useRef<Array<ReturnType<typeof setTimeout>>>([])

  useEffect(() => {
    return () => {
      authoritativeRefreshTimerRefs.current.forEach((timer) => clearTimeout(timer))
      authoritativeRefreshTimerRefs.current = []
    }
  }, [])

  const runSerializedSyncTask = useCallback(async (task: () => Promise<unknown>): Promise<unknown> => {
    const previousTask = syncQueueRef.current
    let releaseQueue = () => {}
    syncQueueRef.current = new Promise<void>((resolve) => {
      releaseQueue = resolve
    })
    await previousTask.catch(() => {})
    try {
      return await task()
    } finally {
      releaseQueue()
    }
  }, [])

  const runSerializedLocalMutationTask = useCallback(async (task: () => Promise<unknown>): Promise<unknown> => {
    const previousTask = localMutationQueueRef.current
    let releaseQueue = () => {}
    localMutationQueueRef.current = new Promise<void>((resolve) => {
      releaseQueue = resolve
    })
    await previousTask.catch(() => {})
    try {
      return await task()
    } finally {
      releaseQueue()
    }
  }, [])

  const scheduleAuthoritativeRefreshProbes = useCallback(() => {
    if (!runtimeController) {
      return
    }
    AUTHORITATIVE_REFRESH_PROBE_DELAYS_MS.forEach((delayMs) => {
      const timer = setTimeout(() => {
        authoritativeRefreshTimerRefs.current = authoritativeRefreshTimerRefs.current.filter((entry) => entry !== timer)
        void (async () => {
          try {
            await runtimeController.invoke('refreshAuthoritativeEvents')
          } catch (error) {
            reportRuntimeError(error)
          }
        })()
      }, delayMs)
      authoritativeRefreshTimerRefs.current.push(timer)
    })
  }, [reportRuntimeError, runtimeController])

  const listPendingMutations = useCallback(async (): Promise<readonly PendingWorkbookMutation[]> => {
    if (!runtimeController) {
      throw new Error('Workbook runtime is not ready')
    }
    const value = await runtimeController.invoke('listPendingMutations')
    assert(isPendingWorkbookMutationList(value), 'Worker returned an invalid pending workbook mutation list')
    return value
  }, [runtimeController])

  const enqueuePendingMutation = useCallback(
    async (mutation: PendingWorkbookMutationInput): Promise<PendingWorkbookMutation> => {
      if (!runtimeController) {
        throw new Error('Workbook runtime is not ready')
      }
      const value = await runtimeController.invoke('enqueuePendingMutation', mutation)
      assert(isPendingWorkbookMutation(value), 'Worker returned an invalid pending mutation')
      return value
    },
    [runtimeController],
  )

  const runZeroMutation = useCallback(
    async (mutation: PendingWorkbookMutation): Promise<{ ok: true } | { ok: false; retryable: boolean; error: Error }> => {
      try {
        const result = zeroRef.current.mutate(buildZeroWorkbookMutation(documentId, mutation))
        const observerResult = observeZeroMutationResult(result)
        if (!observerResult) {
          return { ok: true }
        }
        const remoteResult = await observerResult
        if (!isMutationErrorResult(remoteResult)) {
          return { ok: true }
        }
        const details =
          remoteResult.error.type === 'app' && remoteResult.error.details !== undefined
            ? ` (${JSON.stringify(remoteResult.error.details)})`
            : ''
        return {
          ok: false,
          retryable: remoteResult.error.type === 'zero',
          error: new Error(`${remoteResult.error.message}${details}`),
        }
      } catch (error) {
        return {
          ok: false,
          retryable: true,
          error: error instanceof Error ? error : new Error(toErrorMessage(error)),
        }
      }
    },
    [documentId, zeroRef],
  )

  const drainPendingMutationsLocked = useCallback(async (): Promise<void> => {
    if (!runtimeController || !canAttemptRemoteSync(connectionStateRef.current)) {
      return
    }

    const drainBatch = async (pendingMutations: readonly PendingWorkbookMutation[], index = 0): Promise<void> => {
      const mutation = pendingMutations[index]
      if (!mutation || !canAttemptRemoteSync(connectionStateRef.current)) {
        return
      }

      if (!isPendingWorkbookMutationReadyForSubmission(mutation)) {
        await drainBatch(pendingMutations, index + 1)
        return
      }

      await runtimeController.invoke('recordPendingMutationAttempt', mutation.id)
      scheduleAuthoritativeRefreshProbes()
      const remoteResult = await runZeroMutation(mutation)
      if (!remoteResult.ok) {
        if (!remoteResult.retryable) {
          await runtimeController.invoke('markPendingMutationFailed', mutation.id, remoteResult.error.message)
          return
        }
        return
      }

      await runtimeController.invoke('markPendingMutationSubmitted', mutation.id)
      await drainBatch(pendingMutations, index + 1)
    }

    await drainBatch(await listPendingMutations())
  }, [connectionStateRef, listPendingMutations, runZeroMutation, runtimeController, scheduleAuthoritativeRefreshProbes])

  const drainPendingMutations = useCallback(async (): Promise<void> => {
    try {
      await runSerializedSyncTask(drainPendingMutationsLocked)
    } catch (error) {
      reportRuntimeError(error)
    }
  }, [drainPendingMutationsLocked, reportRuntimeError, runSerializedSyncTask])

  const invokeMutation = useCallback(
    async (method: WorkbookMutationMethod, ...args: unknown[]): Promise<void> => {
      if (!runtimeController) {
        throw new Error('Workbook runtime is not ready')
      }

      let mutation: PendingWorkbookMutationInput
      switch (method) {
        case 'setCellValue': {
          const [sheetName, address, value] = args
          assert(typeof sheetName === 'string' && typeof address === 'string' && isLiteralInput(value), 'Invalid setCellValue args')
          mutation = { method, args: [sheetName, address, value] }
          break
        }
        case 'setCellFormula': {
          const [sheetName, address, formula] = args
          assert(typeof sheetName === 'string' && typeof address === 'string' && typeof formula === 'string', 'Invalid setCellFormula args')
          mutation = { method, args: [sheetName, address, formula] }
          break
        }
        case 'clearCell': {
          const [sheetName, address] = args
          assert(typeof sheetName === 'string' && typeof address === 'string', 'Invalid clearCell args')
          mutation = { method, args: [sheetName, address] }
          break
        }
        case 'clearRange': {
          const [range] = args
          assert(isCellRangeRef(range), 'Invalid clearRange args')
          mutation = { method, args: [range] }
          break
        }
        case 'renderCommit': {
          const [ops] = args
          assert(isCommitOps(ops), 'Invalid renderCommit args')
          mutation = { method, args: [ops] }
          break
        }
        case 'fillRange':
        case 'copyRange':
        case 'moveRange': {
          const [source, target] = args
          assert(isCellRangeRef(source) && isCellRangeRef(target), `Invalid ${method} args`)
          mutation = { method, args: [source, target] }
          break
        }
        case 'insertRows':
        case 'deleteRows':
        case 'insertColumns':
        case 'deleteColumns': {
          const [sheetName, start, count] = args
          assert(
            isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(start) && isWorkbookStructuralCount(count),
            `Invalid ${method} args`,
          )
          mutation = { method, args: [sheetName, start, count] }
          break
        }
        case 'updateRowMetadata': {
          const [sheetName, startRow, count, height, hidden] = args
          assert(
            isWorkbookSheetName(sheetName) &&
              isWorkbookStructuralIndex(startRow) &&
              isWorkbookStructuralCount(count) &&
              (height === null || isWorkbookStructuralSize(height)) &&
              (hidden === null || typeof hidden === 'boolean'),
            'Invalid updateRowMetadata args',
          )
          mutation = { method, args: [sheetName, startRow, count, height, hidden] }
          break
        }
        case 'updateColumnMetadata': {
          const [sheetName, startCol, count, width, hidden] = args
          assert(
            isWorkbookSheetName(sheetName) &&
              isWorkbookStructuralIndex(startCol) &&
              isWorkbookStructuralCount(count) &&
              (width === null || isWorkbookStructuralSize(width)) &&
              (hidden === null || typeof hidden === 'boolean'),
            'Invalid updateColumnMetadata args',
          )
          mutation = { method, args: [sheetName, startCol, count, width, hidden] }
          break
        }
        case 'setFreezePane': {
          const [sheetName, rows, cols] = args
          assert(
            isWorkbookSheetName(sheetName) && isWorkbookStructuralIndex(rows) && isWorkbookStructuralIndex(cols),
            'Invalid setFreezePane args',
          )
          mutation = { method, args: [sheetName, rows, cols] }
          break
        }
        case 'mergeCells':
        case 'unmergeCells': {
          const [range] = args
          assert(isCellRangeRef(range), `Invalid ${method} args`)
          mutation = { method, args: [range] }
          break
        }
        case 'setRangeStyle': {
          const [range, patch] = args
          assert(isCellRangeRef(range) && isCellStylePatchValue(patch), 'Invalid setRangeStyle args')
          mutation = { method, args: [range, patch] }
          break
        }
        case 'clearRangeStyle': {
          const [range, fields] = args
          assert(isCellRangeRef(range) && (fields === undefined || isCellStyleFieldList(fields)), 'Invalid clearRangeStyle args')
          mutation = { method, args: [range, fields] }
          break
        }
        case 'setRangeNumberFormat': {
          const [range, format] = args
          assert(isCellRangeRef(range) && isCellNumberFormatInputValue(format), 'Invalid setRangeNumberFormat args')
          mutation = { method, args: [range, format] }
          break
        }
        case 'clearRangeNumberFormat': {
          const [range] = args
          assert(isCellRangeRef(range), 'Invalid clearRangeNumberFormat args')
          mutation = { method, args: [range] }
          break
        }
        default:
          throw new Error('Unsupported workbook mutation')
      }

      await runSerializedLocalMutationTask(async () => {
        const rollbackOptimisticCell = applyOptimisticCellMutation(workerHandleRef.current?.viewportStore, mutation)
        try {
          await enqueuePendingMutation(mutation)
          if (canAttemptRemoteSync(connectionStateRef.current)) {
            scheduleAuthoritativeRefreshProbes()
          }
        } catch (error) {
          rollbackOptimisticCell?.()
          throw error
        }
      })
      void (async () => {
        try {
          await runSerializedSyncTask(async () => {
            if (canAttemptRemoteSync(connectionStateRef.current)) {
              await drainPendingMutationsLocked()
            }
          })
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    [
      connectionStateRef,
      drainPendingMutationsLocked,
      enqueuePendingMutation,
      reportRuntimeError,
      runSerializedLocalMutationTask,
      runSerializedSyncTask,
      runtimeController,
      scheduleAuthoritativeRefreshProbes,
      workerHandleRef,
    ],
  )

  const invokeColumnWidthMutation = useCallback(
    async (sheetName: string, columnIndex: number, width: number, options?: ViewportAxisSizeMutationOptions): Promise<void> => {
      const initialViewportStore = workerHandleRef.current?.viewportStore
      const previousWidth = initialViewportStore?.getColumnWidths(sheetName)[columnIndex]
      if (options?.deferLocalApplication) {
        await deferViewportAxisSizeFrame()
      }
      const viewportStore = workerHandleRef.current?.viewportStore === initialViewportStore ? initialViewportStore : undefined
      if (viewportStore) {
        const applyOptimisticWidth = () => {
          viewportStore.setColumnWidth(sheetName, columnIndex, width)
        }
        if (options?.flush) {
          flushSync(applyOptimisticWidth)
        } else {
          applyOptimisticWidth()
        }
      }
      if (options?.deferPersistence) {
        await deferViewportAxisSizeFrame()
      }
      try {
        await invokeMutation('updateColumnMetadata', sheetName, columnIndex, 1, width, null)
      } catch (error) {
        if (viewportStore && viewportStore.getColumnWidths(sheetName)[columnIndex] === width) {
          viewportStore.rollbackColumnWidth(sheetName, columnIndex, previousWidth)
        }
        throw error
      }
    },
    [invokeMutation, workerHandleRef],
  )

  const invokeRowHeightMutation = useCallback(
    async (sheetName: string, rowIndex: number, height: number, options?: ViewportAxisSizeMutationOptions): Promise<void> => {
      const initialViewportStore = workerHandleRef.current?.viewportStore
      const previousHeight = initialViewportStore?.getRowHeights(sheetName)[rowIndex]
      if (options?.deferLocalApplication) {
        await deferViewportAxisSizeFrame()
      }
      const viewportStore = workerHandleRef.current?.viewportStore === initialViewportStore ? initialViewportStore : undefined
      if (viewportStore) {
        const applyOptimisticHeight = () => {
          viewportStore.setRowHeight(sheetName, rowIndex, height)
        }
        if (options?.flush) {
          flushSync(applyOptimisticHeight)
        } else {
          applyOptimisticHeight()
        }
      }
      if (options?.deferPersistence) {
        await deferViewportAxisSizeFrame()
      }
      try {
        await invokeMutation('updateRowMetadata', sheetName, rowIndex, 1, height, null)
      } catch (error) {
        if (viewportStore && viewportStore.getRowHeights(sheetName)[rowIndex] === height) {
          viewportStore.rollbackRowHeight(sheetName, rowIndex, previousHeight)
        }
        throw error
      }
    },
    [invokeMutation, workerHandleRef],
  )

  const invokeColumnVisibilityMutation = useCallback(
    async (sheetName: string, columnIndex: number, hidden: boolean): Promise<void> => {
      const viewportStore = workerHandleRef.current?.viewportStore
      const previousHidden = viewportStore?.getHiddenColumns(sheetName)[columnIndex] === true
      const previousSize = viewportStore?.getColumnSizes(sheetName)[columnIndex] ?? viewportStore?.getColumnWidths(sheetName)[columnIndex]
      const nextSize = previousSize ?? PRODUCT_COLUMN_WIDTH
      if (viewportStore) {
        viewportStore.setColumnHidden(sheetName, columnIndex, hidden, nextSize)
      }
      try {
        await invokeMutation('updateColumnMetadata', sheetName, columnIndex, 1, null, hidden)
      } catch (error) {
        viewportStore?.rollbackColumnHidden(sheetName, columnIndex, {
          hidden: previousHidden,
          size: previousSize,
        })
        throw error
      }
    },
    [invokeMutation, workerHandleRef],
  )

  const invokeRowVisibilityMutation = useCallback(
    async (sheetName: string, rowIndex: number, hidden: boolean): Promise<void> => {
      const viewportStore = workerHandleRef.current?.viewportStore
      const previousHidden = viewportStore?.getHiddenRows(sheetName)[rowIndex] === true
      const previousSize = viewportStore?.getRowSizes(sheetName)[rowIndex] ?? viewportStore?.getRowHeights(sheetName)[rowIndex]
      const nextSize = previousSize ?? PRODUCT_ROW_HEIGHT
      if (viewportStore) {
        viewportStore.setRowHidden(sheetName, rowIndex, hidden, nextSize)
      }
      try {
        await invokeMutation('updateRowMetadata', sheetName, rowIndex, 1, null, hidden)
      } catch (error) {
        viewportStore?.rollbackRowHidden(sheetName, rowIndex, {
          hidden: previousHidden,
          size: previousSize,
        })
        throw error
      }
    },
    [invokeMutation, workerHandleRef],
  )

  useEffect(() => {
    if (!runtimeController || !canAttemptRemoteSync(connectionStateName)) {
      return
    }
    void drainPendingMutations()
  }, [connectionStateName, drainPendingMutations, runtimeController])

  const retryPendingMutation = useCallback(
    async (id: string): Promise<void> => {
      if (!runtimeController) {
        throw new Error('Workbook runtime is not ready')
      }
      await runSerializedLocalMutationTask(() => runtimeController.invoke('retryPendingMutation', id))
      await runSerializedSyncTask(async () => {
        if (canAttemptRemoteSync(connectionStateRef.current)) {
          await drainPendingMutationsLocked()
        }
      })
    },
    [connectionStateRef, drainPendingMutationsLocked, runSerializedLocalMutationTask, runSerializedSyncTask, runtimeController],
  )

  return {
    invokeMutation,
    invokeColumnWidthMutation,
    invokeColumnVisibilityMutation,
    invokeRowHeightMutation,
    invokeRowVisibilityMutation,
    retryPendingMutation,
  }
}
