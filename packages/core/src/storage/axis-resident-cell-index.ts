export interface AxisResidentCellIdentity {
  readonly rowId: string
  readonly colId: string
}

export class AxisResidentCellIndex {
  private readonly byCell = new Map<number, AxisResidentCellIdentity>()
  private readonly byRow = new Map<string, Set<number>>()
  private readonly byColumn = new Map<string, Set<number>>()

  set(cellIndex: number, identity: AxisResidentCellIdentity): void {
    this.delete(cellIndex)
    this.byCell.set(cellIndex, identity)
    addToSetMap(this.byRow, identity.rowId, cellIndex)
    addToSetMap(this.byColumn, identity.colId, cellIndex)
  }

  get(cellIndex: number): AxisResidentCellIdentity | undefined {
    return this.byCell.get(cellIndex)
  }

  delete(cellIndex: number): boolean {
    const existing = this.byCell.get(cellIndex)
    if (!existing) {
      return false
    }
    this.byCell.delete(cellIndex)
    deleteFromSetMap(this.byRow, existing.rowId, cellIndex)
    deleteFromSetMap(this.byColumn, existing.colId, cellIndex)
    return true
  }

  clear(): void {
    this.byCell.clear()
    this.byRow.clear()
    this.byColumn.clear()
  }

  cellsInRow(rowId: string): number[] {
    return sortedCells(this.byRow.get(rowId))
  }

  cellsInColumn(colId: string): number[] {
    return sortedCells(this.byColumn.get(colId))
  }

  cellsInRows(rowIds: readonly string[]): number[] {
    return sortedUniqueCells(rowIds.flatMap((rowId) => this.cellsInRow(rowId)))
  }

  cellsInColumns(colIds: readonly string[]): number[] {
    return sortedUniqueCells(colIds.flatMap((colId) => this.cellsInColumn(colId)))
  }

  cellsInRowsUnordered(rowIds: readonly string[]): number[] {
    return uniqueCells(rowIds.flatMap((rowId) => [...(this.byRow.get(rowId) ?? [])]))
  }

  cellsInColumnsUnordered(colIds: readonly string[]): number[] {
    return uniqueCells(colIds.flatMap((colId) => [...(this.byColumn.get(colId) ?? [])]))
  }
}

function addToSetMap(map: Map<string, Set<number>>, key: string, value: number): void {
  let values = map.get(key)
  if (!values) {
    values = new Set<number>()
    map.set(key, values)
  }
  values.add(value)
}

function deleteFromSetMap(map: Map<string, Set<number>>, key: string, value: number): void {
  const values = map.get(key)
  if (!values) {
    return
  }
  values.delete(value)
  if (values.size === 0) {
    map.delete(key)
  }
}

function sortedCells(values: ReadonlySet<number> | undefined): number[] {
  return values ? [...values].toSorted((left, right) => left - right) : []
}

function sortedUniqueCells(values: readonly number[]): number[] {
  return [...new Set(values)].toSorted((left, right) => left - right)
}

function uniqueCells(values: readonly number[]): number[] {
  return [...new Set(values)]
}
