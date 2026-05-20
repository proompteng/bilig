const initialStyleIndexCapacity = 256

interface StyleRun {
  row: number
  startColumn: number
  endColumn: number
  styleIndex: number
}

interface StyleRect {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  styleIndex: number
}

export class ImportedWorksheetStyleIndexArena {
  private rectStartRows: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private rectEndRows: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private rectStartColumns: Uint16Array<ArrayBuffer> = new Uint16Array(initialStyleIndexCapacity)
  private rectEndColumns: Uint16Array<ArrayBuffer> = new Uint16Array(initialStyleIndexCapacity)
  private rectStyleIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private rectLength = 0
  private rows: Uint32Array<ArrayBuffer> | undefined
  private columns: Uint16Array<ArrayBuffer> | undefined
  private styleIndexes: Uint32Array<ArrayBuffer> | undefined
  private length = 0
  private rowMajorOrdered = true
  private coordinatesDeferred = false
  private lastRow = -1
  private lastColumn = -1
  private activeRun: StyleRun | undefined
  private readonly openRects = new Map<string, StyleRect>()
  private requiredStyleIndexes: Set<number> | undefined

  get count(): number {
    return this.length
  }

  get hasCoordinateStorage(): boolean {
    return !this.coordinatesDeferred
  }

  get isRowMajorOrdered(): boolean {
    return this.rowMajorOrdered
  }

  add(row: number, column: number, styleIndex: number): void {
    if (!this.rowMajorOrdered) {
      this.appendExpanded(row, column, styleIndex)
      return
    }
    if (row < this.lastRow || (row === this.lastRow && column < this.lastColumn)) {
      this.expandCompressedStorage()
      this.rowMajorOrdered = false
      this.appendExpanded(row, column, styleIndex)
      return
    }
    this.addRowMajor(row, column, styleIndex)
    this.length += 1
    this.lastRow = row
    this.lastColumn = column
  }

  addRequiredStyleIndex(styleIndex: number): void {
    if (this.length > 0 || this.rectLength > 0 || this.activeRun || this.openRects.size > 0) {
      this.requiredStyleIndexes ??= new Set()
      this.requiredStyleIndexes.add(styleIndex)
      return
    }
    this.coordinatesDeferred = true
    this.requiredStyleIndexes ??= new Set()
    this.requiredStyleIndexes.add(styleIndex)
  }

  collectRequiredStyleIndexes(output: Set<number>): void {
    for (const styleIndex of this.requiredStyleIndexes ?? []) {
      output.add(styleIndex)
    }
    this.finishCompressedStorage()
    if (!this.rowMajorOrdered) {
      for (let index = 0; index < this.length; index += 1) {
        output.add(this.styleIndexes?.[index] ?? 0)
      }
      return
    }
    for (let index = 0; index < this.rectLength; index += 1) {
      output.add(this.rectStyleIndexes[index] ?? 0)
    }
  }

  forEach(callback: (row: number, column: number, styleIndex: number) => void): void {
    this.finishCompressedStorage()
    if (!this.rowMajorOrdered) {
      for (let index = 0; index < this.length; index += 1) {
        callback(this.rows?.[index] ?? 0, this.columns?.[index] ?? 0, this.styleIndexes?.[index] ?? 0)
      }
      return
    }
    for (const rect of this.sortedRects()) {
      for (let row = rect.startRow; row <= rect.endRow; row += 1) {
        for (let column = rect.startColumn; column <= rect.endColumn; column += 1) {
          callback(row, column, rect.styleIndex)
        }
      }
    }
  }

  forEachCompressedRange(
    callback: (startRow: number, endRow: number, startColumn: number, endColumn: number, styleIndex: number) => void,
  ): boolean {
    this.finishCompressedStorage()
    if (!this.rowMajorOrdered) {
      return false
    }
    for (const rect of this.sortedRects()) {
      callback(rect.startRow, rect.endRow, rect.startColumn, rect.endColumn, rect.styleIndex)
    }
    return true
  }

  release(): void {
    this.rectStartRows = new Uint32Array(0)
    this.rectEndRows = new Uint32Array(0)
    this.rectStartColumns = new Uint16Array(0)
    this.rectEndColumns = new Uint16Array(0)
    this.rectStyleIndexes = new Uint32Array(0)
    this.rectLength = 0
    this.rows = undefined
    this.columns = undefined
    this.styleIndexes = undefined
    this.length = 0
    this.rowMajorOrdered = true
    this.coordinatesDeferred = false
    this.lastRow = -1
    this.lastColumn = -1
    this.activeRun = undefined
    this.openRects.clear()
    this.requiredStyleIndexes = undefined
  }

  private addRowMajor(row: number, column: number, styleIndex: number): void {
    if (this.activeRun && this.activeRun.row !== row) {
      const previousRow = this.activeRun.row
      this.flushActiveRun()
      this.flushOpenRectsNotEndingAt(previousRow)
    }
    if (
      this.activeRun &&
      this.activeRun.row === row &&
      this.activeRun.endColumn + 1 === column &&
      this.activeRun.styleIndex === styleIndex
    ) {
      this.activeRun.endColumn = column
      return
    }
    this.flushActiveRun()
    this.activeRun = { row, startColumn: column, endColumn: column, styleIndex }
  }

  private appendExpanded(row: number, column: number, styleIndex: number): void {
    this.ensureExpandedCapacity(this.length + 1)
    this.rows![this.length] = row
    this.columns![this.length] = column
    this.styleIndexes![this.length] = styleIndex
    this.length += 1
    this.lastRow = row
    this.lastColumn = column
  }

  private flushActiveRun(): void {
    if (!this.activeRun) {
      return
    }
    this.appendRunToOpenRect(this.activeRun)
    this.activeRun = undefined
  }

  private appendRunToOpenRect(run: StyleRun): void {
    const key = rectKey(run.startColumn, run.endColumn, run.styleIndex)
    const rect = this.openRects.get(key)
    if (rect && rect.endRow === run.row - 1) {
      rect.endRow = run.row
      return
    }
    if (rect) {
      this.flushOpenRect(key)
    }
    this.openRects.set(key, {
      startRow: run.row,
      endRow: run.row,
      startColumn: run.startColumn,
      endColumn: run.endColumn,
      styleIndex: run.styleIndex,
    })
  }

  private flushOpenRectsNotEndingAt(row: number): void {
    for (const [key, rect] of this.openRects) {
      if (rect.endRow !== row) {
        this.flushOpenRect(key)
      }
    }
  }

  private finishCompressedStorage(): void {
    this.flushActiveRun()
    while (this.openRects.size > 0) {
      const key = this.openRects.keys().next().value
      if (key === undefined) {
        return
      }
      this.flushOpenRect(key)
    }
  }

  private flushOpenRect(key: string): void {
    const rect = this.openRects.get(key)
    if (!rect) {
      return
    }
    this.ensureRectCapacity(this.rectLength + 1)
    this.rectStartRows[this.rectLength] = rect.startRow
    this.rectEndRows[this.rectLength] = rect.endRow
    this.rectStartColumns[this.rectLength] = rect.startColumn
    this.rectEndColumns[this.rectLength] = rect.endColumn
    this.rectStyleIndexes[this.rectLength] = rect.styleIndex
    this.rectLength += 1
    this.openRects.delete(key)
  }

  private expandCompressedStorage(): void {
    this.finishCompressedStorage()
    this.rows = new Uint32Array(Math.max(initialStyleIndexCapacity, this.length))
    this.columns = new Uint16Array(this.rows.length)
    this.styleIndexes = new Uint32Array(this.rows.length)
    let outputIndex = 0
    for (const rect of this.sortedRects()) {
      for (let row = rect.startRow; row <= rect.endRow; row += 1) {
        for (let column = rect.startColumn; column <= rect.endColumn; column += 1) {
          this.rows[outputIndex] = row
          this.columns[outputIndex] = column
          this.styleIndexes[outputIndex] = rect.styleIndex
          outputIndex += 1
        }
      }
    }
    this.rectStartRows = new Uint32Array(0)
    this.rectEndRows = new Uint32Array(0)
    this.rectStartColumns = new Uint16Array(0)
    this.rectEndColumns = new Uint16Array(0)
    this.rectStyleIndexes = new Uint32Array(0)
    this.rectLength = 0
  }

  private sortedRects(): StyleRect[] {
    const rects: StyleRect[] = []
    for (let index = 0; index < this.rectLength; index += 1) {
      rects.push({
        startRow: this.rectStartRows[index] ?? 0,
        endRow: this.rectEndRows[index] ?? 0,
        startColumn: this.rectStartColumns[index] ?? 0,
        endColumn: this.rectEndColumns[index] ?? 0,
        styleIndex: this.rectStyleIndexes[index] ?? 0,
      })
    }
    return rects.toSorted(
      (left, right) =>
        left.startRow - right.startRow ||
        left.startColumn - right.startColumn ||
        left.endRow - right.endRow ||
        left.endColumn - right.endColumn ||
        left.styleIndex - right.styleIndex,
    )
  }

  private ensureRectCapacity(nextLength: number): void {
    if (nextLength <= this.rectStartRows.length) {
      return
    }
    let nextCapacity = this.rectStartRows.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.rectStartRows = growUint32Array(this.rectStartRows, nextCapacity)
    this.rectEndRows = growUint32Array(this.rectEndRows, nextCapacity)
    this.rectStartColumns = growUint16Array(this.rectStartColumns, nextCapacity)
    this.rectEndColumns = growUint16Array(this.rectEndColumns, nextCapacity)
    this.rectStyleIndexes = growUint32Array(this.rectStyleIndexes, nextCapacity)
  }

  private ensureExpandedCapacity(nextLength: number): void {
    if (!this.rows || !this.columns || !this.styleIndexes) {
      this.rows = new Uint32Array(initialStyleIndexCapacity)
      this.columns = new Uint16Array(initialStyleIndexCapacity)
      this.styleIndexes = new Uint32Array(initialStyleIndexCapacity)
    }
    if (nextLength <= this.rows.length) {
      return
    }
    let nextCapacity = this.rows.length
    while (nextCapacity < nextLength) {
      nextCapacity *= 2
    }
    this.rows = growUint32Array(this.rows, nextCapacity)
    this.columns = growUint16Array(this.columns, nextCapacity)
    this.styleIndexes = growUint32Array(this.styleIndexes, nextCapacity)
  }
}

function rectKey(startColumn: number, endColumn: number, styleIndex: number): string {
  return `${String(startColumn)}\t${String(endColumn)}\t${String(styleIndex)}`
}

function growUint32Array(source: Uint32Array<ArrayBuffer>, nextCapacity: number): Uint32Array<ArrayBuffer> {
  const output = new Uint32Array(nextCapacity)
  output.set(source.subarray(0, Math.min(source.length, nextCapacity)))
  return output
}

function growUint16Array(source: Uint16Array<ArrayBuffer>, nextCapacity: number): Uint16Array<ArrayBuffer> {
  const output = new Uint16Array(nextCapacity)
  output.set(source.subarray(0, Math.min(source.length, nextCapacity)))
  return output
}
