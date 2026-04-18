export interface LogicalCellLocation {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

export class CellPageStore {
  constructor(
    private readonly cells: Map<string, number>,
    private readonly keyForLocation: (location: LogicalCellLocation) => string,
  ) {}

  key(location: LogicalCellLocation): string {
    return this.keyForLocation(location)
  }

  get(location: LogicalCellLocation): number | undefined {
    return this.cells.get(this.key(location))
  }

  set(location: LogicalCellLocation, cellIndex: number): void {
    this.cells.set(this.key(location), cellIndex)
  }

  delete(location: LogicalCellLocation): boolean {
    return this.cells.delete(this.key(location))
  }
}
