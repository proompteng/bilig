export interface CellAxisIdentity {
  readonly sheetId: number
  readonly rowId: string
  readonly colId: string
}

export class CellAxisIdentityStore {
  private readonly identities = new Map<number, CellAxisIdentity>()

  get(cellIndex: number): CellAxisIdentity | undefined {
    return this.identities.get(cellIndex)
  }

  set(cellIndex: number, identity: CellAxisIdentity): void {
    this.identities.set(cellIndex, identity)
  }

  delete(cellIndex: number): boolean {
    return this.identities.delete(cellIndex)
  }

  clear(): void {
    this.identities.clear()
  }

  forEach(callback: (identity: CellAxisIdentity, cellIndex: number) => void): void {
    this.identities.forEach((identity, cellIndex) => {
      callback(identity, cellIndex)
    })
  }

  entries(): Array<readonly [number, CellAxisIdentity]> {
    return [...this.identities.entries()]
  }
}
