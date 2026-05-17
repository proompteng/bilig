export interface LogicalCellLocation {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

export type CellPageStoreRebuildSource = (callback: (location: LogicalCellLocation, cellIndex: number) => void) => void
export type LogicalCellPartsKeyFactory = (sheetId: number, rowId: string, colId: string) => string

export class CellPageStore {
  private pagesDirty = false
  private readonly keyForParts: LogicalCellPartsKeyFactory

  constructor(
    private readonly cells: Map<string, number>,
    private readonly keyForLocation: (location: LogicalCellLocation) => string,
    private readonly rebuildSource?: CellPageStoreRebuildSource,
    keyForParts?: LogicalCellPartsKeyFactory,
  ) {
    this.keyForParts = keyForParts ?? ((sheetId, rowId, colId) => this.keyForLocation({ sheetId, rowId, colId }))
  }

  key(location: LogicalCellLocation): string {
    return this.keyForLocation(location)
  }

  get(location: LogicalCellLocation): number | undefined {
    this.ensurePages()
    return this.cells.get(this.key(location))
  }

  set(location: LogicalCellLocation, cellIndex: number): void {
    this.ensurePages()
    this.cells.set(this.key(location), cellIndex)
  }

  setDeferred(location: LogicalCellLocation, cellIndex: number): void {
    this.cells.set(this.key(location), cellIndex)
  }

  setDeferredParts(sheetId: number, rowId: string, colId: string, cellIndex: number): void {
    this.cells.set(this.keyForParts(sheetId, rowId, colId), cellIndex)
  }

  deferRebuild(): void {
    this.pagesDirty = true
  }

  delete(location: LogicalCellLocation): boolean {
    this.ensurePages()
    return this.cells.delete(this.key(location))
  }

  private ensurePages(): void {
    if (!this.pagesDirty) {
      return
    }
    this.cells.clear()
    this.rebuildSource?.((location, cellIndex) => {
      this.cells.set(this.key(location), cellIndex)
    })
    this.pagesDirty = false
  }
}
