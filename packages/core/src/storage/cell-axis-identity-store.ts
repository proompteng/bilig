export interface CellAxisIdentity {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

export class CellAxisIdentityStore {
  private readonly sheetIds: Array<number | undefined> = []
  private readonly rowIds: Array<string | undefined> = []
  private readonly colIds: Array<string | undefined> = []

  get(cellIndex: number): CellAxisIdentity | undefined {
    const sheetId = this.sheetIds[cellIndex]
    const rowId = this.rowIds[cellIndex]
    const colId = this.colIds[cellIndex]
    return sheetId === undefined || rowId === undefined || colId === undefined
      ? undefined
      : {
          sheetId,
          rowId,
          colId,
        }
  }

  set(cellIndex: number, identity: CellAxisIdentity): void {
    this.setParts(cellIndex, identity.sheetId, identity.rowId, identity.colId)
  }

  setParts(cellIndex: number, sheetId: number, rowId: string, colId: string): void {
    this.sheetIds[cellIndex] = sheetId
    this.rowIds[cellIndex] = rowId
    this.colIds[cellIndex] = colId
  }

  delete(cellIndex: number): boolean {
    if (this.sheetIds[cellIndex] === undefined || this.rowIds[cellIndex] === undefined || this.colIds[cellIndex] === undefined) {
      return false
    }
    this.sheetIds[cellIndex] = undefined
    this.rowIds[cellIndex] = undefined
    this.colIds[cellIndex] = undefined
    return true
  }

  clear(): void {
    this.sheetIds.length = 0
    this.rowIds.length = 0
    this.colIds.length = 0
  }

  forEach(callback: (identity: CellAxisIdentity, cellIndex: number) => void): void {
    for (let cellIndex = 0; cellIndex < this.rowIds.length; cellIndex += 1) {
      const sheetId = this.sheetIds[cellIndex]
      const rowId = this.rowIds[cellIndex]
      const colId = this.colIds[cellIndex]
      if (sheetId === undefined || rowId === undefined || colId === undefined) {
        continue
      }
      callback({ sheetId, rowId, colId }, cellIndex)
    }
  }

  entries(): Array<readonly [number, CellAxisIdentity]> {
    const entries: Array<readonly [number, CellAxisIdentity]> = []
    this.forEach((identity, cellIndex) => {
      entries.push([cellIndex, identity])
    })
    return entries
  }
}
