import type { WorkPaperChange } from './work-paper-types.js'

type SheetOrder = readonly { id: number; order: number }[]

export function orderWorkPaperCellChanges(
  changes: WorkPaperChange[],
  sheets: SheetOrder,
  explicitChangedCount?: number,
): WorkPaperChange[] {
  if (changes.length < 2) {
    return changes
  }

  const compare = sheets.length === 1 ? compareSingleSheetCellChanges : compareWorkPaperCellChanges(sheets)
  const split = explicitChangedCount === undefined ? undefined : Math.min(explicitChangedCount, changes.length)

  if (
    split !== undefined &&
    split > 0 &&
    split < changes.length &&
    isSortedCellChangeSlice(changes, compare, 0, split) &&
    isSortedCellChangeSlice(changes, compare, split, changes.length)
  ) {
    return mergeSortedCellChangeSlices(changes, compare, split)
  }

  if (isSortedCellChangeSlice(changes, compare, 0, changes.length)) {
    return changes
  }

  return changes.toSorted(compare)
}

function compareSingleSheetCellChanges(left: WorkPaperChange, right: WorkPaperChange): number {
  if (left.kind !== 'cell' || right.kind !== 'cell') {
    return 0
  }
  return left.address.row - right.address.row || left.address.col - right.address.col
}

function compareWorkPaperCellChanges(sheets: SheetOrder): (left: WorkPaperChange, right: WorkPaperChange) => number {
  const orderBySheet = new Map(sheets.map((sheet) => [sheet.id, sheet.order]))
  return (left, right) => {
    if (left.kind !== 'cell' || right.kind !== 'cell') {
      return 0
    }
    return (
      (orderBySheet.get(left.address.sheet) ?? 0) - (orderBySheet.get(right.address.sheet) ?? 0) ||
      left.address.row - right.address.row ||
      left.address.col - right.address.col
    )
  }
}

function isSortedCellChangeSlice(
  changes: readonly WorkPaperChange[],
  compare: (left: WorkPaperChange, right: WorkPaperChange) => number,
  start: number,
  end: number,
): boolean {
  for (let index = start + 1; index < end; index += 1) {
    if (compare(changes[index - 1]!, changes[index]!) > 0) {
      return false
    }
  }
  return true
}

function mergeSortedCellChangeSlices(
  changes: readonly WorkPaperChange[],
  compare: (left: WorkPaperChange, right: WorkPaperChange) => number,
  split: number,
): WorkPaperChange[] {
  const merged: WorkPaperChange[] = []
  let leftIndex = 0
  let rightIndex = split
  while (leftIndex < split && rightIndex < changes.length) {
    const left = changes[leftIndex]!
    const right = changes[rightIndex]!
    if (compare(left, right) <= 0) {
      merged.push(left)
      leftIndex += 1
    } else {
      merged.push(right)
      rightIndex += 1
    }
  }
  while (leftIndex < split) {
    merged.push(changes[leftIndex]!)
    leftIndex += 1
  }
  while (rightIndex < changes.length) {
    merged.push(changes[rightIndex]!)
    rightIndex += 1
  }
  return merged
}
