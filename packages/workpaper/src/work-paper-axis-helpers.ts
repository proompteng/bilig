import { WorkPaperInvalidArgumentsError, WorkPaperOperationError } from './work-paper-errors.js'
import type { WorkPaperAxisInterval, WorkPaperAxisSwapMapping, WorkPaperChange } from './work-paper-types.js'

export type WorkPaperAxisKind = 'row' | 'column'
export type WorkPaperAxisIntervalEditMode = 'add' | 'remove'
export type WorkPaperAxisMoveCallback = (start: number, count: number, target: number) => void

export interface WorkPaperAxisEditRuntime {
  canSwapAxisIndexes(axis: WorkPaperAxisKind, sheetId: number, mappings: readonly WorkPaperAxisSwapMapping[]): boolean
  canSetAxisOrder(axis: WorkPaperAxisKind, sheetId: number, order: readonly number[]): boolean
  canEditAxisIntervals(
    axis: WorkPaperAxisKind,
    mode: WorkPaperAxisIntervalEditMode,
    sheetId: number,
    indexes: readonly WorkPaperAxisInterval[],
  ): boolean
  canMoveAxis(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): boolean
  canUseTrackedStructuralFastPath(): boolean
  isTrackedBatchFastPathActive(): boolean
  batch(operations: () => void): WorkPaperChange[]
  batchStructuralChanges(operations: () => void): WorkPaperChange[]
  captureAxisChange(operations: () => void): WorkPaperChange[]
  captureTrackedStructuralChanges(operations: () => void): WorkPaperChange[]
  moveAxis(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): WorkPaperChange[]
  applyAxisIntervalEdit(axis: WorkPaperAxisKind, mode: WorkPaperAxisIntervalEditMode, sheetId: number, start: number, amount: number): void
  applyAxisMove(axis: WorkPaperAxisKind, sheetId: number, start: number, count: number, target: number): void
}

export function normalizeAxisIntervals(
  startOrInterval: number | WorkPaperAxisInterval,
  countOrInterval?: number | WorkPaperAxisInterval,
  restIntervals: readonly WorkPaperAxisInterval[] = [],
): Array<[number, number]> {
  if (typeof startOrInterval === 'number') {
    if (Array.isArray(countOrInterval)) {
      throw new WorkPaperInvalidArgumentsError('Axis interval count must be a number')
    }
    const resolvedCount = typeof countOrInterval === 'number' ? countOrInterval : 1
    return [[startOrInterval, resolvedCount]]
  }
  if (typeof countOrInterval === 'number') {
    throw new WorkPaperInvalidArgumentsError('Axis interval count is only valid with a numeric start')
  }
  return [startOrInterval, ...(countOrInterval ? [countOrInterval] : []), ...restIntervals].map(
    ([start, count]) => [start, count ?? 1] as [number, number],
  )
}

export function normalizeAxisSwapMappings(
  label: 'row' | 'column',
  startOrMappings: number | readonly WorkPaperAxisSwapMapping[],
  end?: number,
): WorkPaperAxisSwapMapping[] {
  if (typeof startOrMappings === 'number') {
    if (end === undefined) {
      throw new WorkPaperInvalidArgumentsError(`${label} swap requires two indexes`)
    }
    return [[startOrMappings, end]]
  }
  return [...startOrMappings]
}

export function applyWorkPaperAxisSwapMappings(mappings: readonly WorkPaperAxisSwapMapping[], moveAxis: WorkPaperAxisMoveCallback): void {
  mappings.forEach(([from, to]) => {
    if (from === to) {
      return
    }
    if (from < to) {
      moveAxis(from, 1, to)
      moveAxis(to - 1, 1, from)
      return
    }
    moveAxis(from, 1, to)
    moveAxis(to + 1, 1, from)
  })
}

export function applyWorkPaperAxisOrder(axisOrder: readonly number[], moveAxis: WorkPaperAxisMoveCallback): void {
  const current = axisOrder.toSorted((left, right) => left - right)
  axisOrder.forEach((targetOriginalIndex, targetIndex) => {
    const currentIndex = current.indexOf(targetOriginalIndex)
    if (currentIndex === targetIndex) {
      return
    }
    moveAxis(currentIndex, 1, targetIndex)
    const [moved] = current.splice(currentIndex, 1)
    current.splice(targetIndex, 0, moved!)
  })
}

export function swapWorkPaperAxisIndexes(
  runtime: WorkPaperAxisEditRuntime,
  axis: WorkPaperAxisKind,
  sheetId: number,
  firstOrMappings: number | readonly WorkPaperAxisSwapMapping[],
  second?: number,
): WorkPaperChange[] {
  const mappings = normalizeAxisSwapMappings(axis, firstOrMappings, second)
  if (!runtime.canSwapAxisIndexes(axis, sheetId, mappings)) {
    throw new WorkPaperOperationError(`${axisLabel(axis, 'plural')} cannot be swapped`)
  }
  return runtime.batch(() => {
    applyWorkPaperAxisSwapMappings(mappings, (start, count, target) => {
      runtime.moveAxis(axis, sheetId, start, count, target)
    })
  })
}

export function setWorkPaperAxisOrder(
  runtime: WorkPaperAxisEditRuntime,
  axis: WorkPaperAxisKind,
  sheetId: number,
  order: readonly number[],
): WorkPaperChange[] {
  if (!runtime.canSetAxisOrder(axis, sheetId, order)) {
    throw new WorkPaperOperationError(`${axisLabel(axis, 'singular')} order is invalid`)
  }
  return runtime.batch(() => {
    applyWorkPaperAxisOrder(order, (start, count, target) => {
      runtime.moveAxis(axis, sheetId, start, count, target)
    })
  })
}

export function editWorkPaperAxisIntervals(
  runtime: WorkPaperAxisEditRuntime,
  axis: WorkPaperAxisKind,
  mode: WorkPaperAxisIntervalEditMode,
  sheetId: number,
  startOrInterval: number | WorkPaperAxisInterval,
  countOrInterval: number | WorkPaperAxisInterval | undefined,
  restIntervals: readonly WorkPaperAxisInterval[],
): WorkPaperChange[] {
  const indexes = normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals)
  if (!runtime.canEditAxisIntervals(axis, mode, sheetId, indexes)) {
    throw new WorkPaperOperationError(`${axisLabel(axis, 'plural')} cannot be ${mode === 'add' ? 'added' : 'removed'}`)
  }
  const orderedIndexes = mode === 'remove' ? indexes.toSorted((left, right) => right[0] - left[0]) : indexes
  if (indexes.length === 1 && runtime.canUseTrackedStructuralFastPath()) {
    const [start, amount] = indexes[0]!
    return runtime.captureTrackedStructuralChanges(() => {
      runtime.applyAxisIntervalEdit(axis, mode, sheetId, start, amount)
    })
  }
  if (indexes.length === 1 && runtime.isTrackedBatchFastPathActive()) {
    const [start, amount] = indexes[0]!
    runtime.applyAxisIntervalEdit(axis, mode, sheetId, start, amount)
    return []
  }
  return runtime.batchStructuralChanges(() => {
    orderedIndexes.forEach(([start, amount]) => {
      runtime.applyAxisIntervalEdit(axis, mode, sheetId, start, amount)
    })
  })
}

export function moveWorkPaperAxis(
  runtime: WorkPaperAxisEditRuntime,
  axis: WorkPaperAxisKind,
  sheetId: number,
  start: number,
  count: number,
  target: number,
): WorkPaperChange[] {
  if (!runtime.canMoveAxis(axis, sheetId, start, count, target)) {
    throw new WorkPaperOperationError(`${axisLabel(axis, 'plural')} cannot be moved`)
  }
  const move = () => {
    runtime.applyAxisMove(axis, sheetId, start, count, target)
  }
  return runtime.canUseTrackedStructuralFastPath() ? runtime.captureTrackedStructuralChanges(move) : runtime.captureAxisChange(move)
}

function axisLabel(axis: WorkPaperAxisKind, count: 'singular' | 'plural'): string {
  if (axis === 'row') {
    return count === 'singular' ? 'Row' : 'Rows'
  }
  return count === 'singular' ? 'Column' : 'Columns'
}
