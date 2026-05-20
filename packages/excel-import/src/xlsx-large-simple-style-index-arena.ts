const initialStyleIndexCapacity = 256

export class ImportedWorksheetStyleIndexArena {
  private rows: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private columns: Uint16Array<ArrayBuffer> = new Uint16Array(initialStyleIndexCapacity)
  private styleIndexes: Uint32Array<ArrayBuffer> = new Uint32Array(initialStyleIndexCapacity)
  private length = 0
  private rowMajorOrdered = true
  private lastRow = -1
  private lastColumn = -1

  get count(): number {
    return this.length
  }

  get isRowMajorOrdered(): boolean {
    return this.rowMajorOrdered
  }

  add(row: number, column: number, styleIndex: number): void {
    this.ensureCapacity(this.length + 1)
    const index = this.length
    this.length += 1
    if (row < this.lastRow || (row === this.lastRow && column < this.lastColumn)) {
      this.rowMajorOrdered = false
    }
    this.lastRow = row
    this.lastColumn = column
    this.rows[index] = row
    this.columns[index] = column
    this.styleIndexes[index] = styleIndex
  }

  collectRequiredStyleIndexes(output: Set<number>): void {
    for (let index = 0; index < this.length; index += 1) {
      output.add(this.styleIndexes[index] ?? 0)
    }
  }

  forEach(callback: (row: number, column: number, styleIndex: number) => void): void {
    for (let index = 0; index < this.length; index += 1) {
      callback(this.rows[index] ?? 0, this.columns[index] ?? 0, this.styleIndexes[index] ?? 0)
    }
  }

  release(): void {
    this.rows = new Uint32Array(0)
    this.columns = new Uint16Array(0)
    this.styleIndexes = new Uint32Array(0)
    this.length = 0
    this.rowMajorOrdered = true
    this.lastRow = -1
    this.lastColumn = -1
  }

  private ensureCapacity(nextLength: number): void {
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
