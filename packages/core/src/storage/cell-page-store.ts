export interface VisibleCellLocation {
  readonly sheetId: number
  readonly row: number
  readonly col: number
}

export class CellPageStore {
  constructor(
    private readonly cells: Map<number, number>,
    private readonly keyForLocation: (location: VisibleCellLocation) => number,
  ) {}

  key(location: VisibleCellLocation): number {
    return this.keyForLocation(location)
  }

  get(location: VisibleCellLocation): number | undefined {
    return this.cells.get(this.key(location))
  }

  set(location: VisibleCellLocation, cellIndex: number): void {
    this.cells.set(this.key(location), cellIndex)
  }

  delete(location: VisibleCellLocation): boolean {
    return this.cells.delete(this.key(location))
  }
}
