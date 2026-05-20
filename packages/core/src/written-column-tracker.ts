export interface WrittenColumnTracker {
  smallColumns: number
  columns?: Uint8Array
  count: number
}

export function createWrittenColumnTracker(): WrittenColumnTracker {
  return {
    smallColumns: 0,
    count: 0,
  }
}

export function markWrittenColumn(tracker: WrittenColumnTracker, col: number): void {
  if (col < 30) {
    const bit = 1 << col
    if ((tracker.smallColumns & bit) !== 0) {
      return
    }
    tracker.smallColumns |= bit
    tracker.count += 1
    return
  }
  let columns = tracker.columns
  if (!columns) {
    columns = new Uint8Array(Math.max(32, col + 1))
    tracker.columns = columns
  } else if (col >= columns.length) {
    let nextLength = columns.length
    while (nextLength <= col) {
      nextLength *= 2
    }
    const nextColumns = new Uint8Array(nextLength)
    nextColumns.set(columns)
    columns = nextColumns
    tracker.columns = columns
  }
  if (columns[col] !== 0) {
    return
  }
  columns[col] = 1
  tracker.count += 1
}

export function materializeWrittenColumns(tracker: WrittenColumnTracker): Uint32Array {
  const columns = new Uint32Array(tracker.count)
  let writeIndex = 0
  let smallColumns = tracker.smallColumns
  while (smallColumns !== 0) {
    const bit = smallColumns & -smallColumns
    columns[writeIndex] = 31 - Math.clz32(bit)
    writeIndex += 1
    smallColumns &= smallColumns - 1
  }
  const largeColumns = tracker.columns
  if (largeColumns) {
    for (let col = 30; col < largeColumns.length; col += 1) {
      if (largeColumns[col] === 0) {
        continue
      }
      columns[writeIndex] = col
      writeIndex += 1
    }
  }
  return columns
}
