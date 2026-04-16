import type { CellStore } from './cell-store.js'

export class FormulaTable<T extends { cellIndex: number }> {
  private readonly records: Array<T | undefined> = []
  private readonly freeSlots: number[] = []
  private activeCount = 0

  constructor(private readonly store: CellStore) {}

  get size(): number {
    return this.activeCount
  }

  get(cellIndex: number): T | undefined {
    const formulaId = this.store.formulaIds[cellIndex] ?? 0
    return formulaId === 0 ? undefined : this.records[formulaId - 1]
  }

  has(cellIndex: number): boolean {
    return this.get(cellIndex) !== undefined
  }

  set(cellIndex: number, record: T): number {
    const existingId = this.store.formulaIds[cellIndex] ?? 0
    if (existingId !== 0) {
      this.records[existingId - 1] = record
      return existingId
    }

    const slot = this.freeSlots.pop() ?? this.records.length
    this.records[slot] = record
    this.store.formulaIds[cellIndex] = slot + 1
    this.activeCount += 1
    return slot + 1
  }

  delete(cellIndex: number): T | undefined {
    const formulaId = this.store.formulaIds[cellIndex] ?? 0
    if (formulaId === 0) {
      return undefined
    }

    const slot = formulaId - 1
    const existing = this.records[slot]
    if (existing !== undefined) {
      this.records[slot] = undefined
      this.freeSlots.push(slot)
      this.activeCount -= 1
    }
    this.store.formulaIds[cellIndex] = 0
    return existing
  }

  clear(): void {
    for (let index = 0; index < this.records.length; index += 1) {
      const record = this.records[index]
      if (record !== undefined) {
        this.store.formulaIds[record.cellIndex] = 0
      }
    }
    this.records.length = 0
    this.freeSlots.length = 0
    this.activeCount = 0
  }

  forEach(callback: (record: T, cellIndex: number) => void): void {
    for (let index = 0; index < this.records.length; index += 1) {
      const record = this.records[index]
      if (record !== undefined) {
        callback(record, record.cellIndex)
      }
    }
  }

  *keys(): IterableIterator<number> {
    for (let index = 0; index < this.records.length; index += 1) {
      const record = this.records[index]
      if (record !== undefined) {
        yield record.cellIndex
      }
    }
  }

  *values(): IterableIterator<T> {
    for (let index = 0; index < this.records.length; index += 1) {
      const record = this.records[index]
      if (record !== undefined) {
        yield record
      }
    }
  }

  *entries(): IterableIterator<[number, T]> {
    for (let index = 0; index < this.records.length; index += 1) {
      const record = this.records[index]
      if (record !== undefined) {
        yield [record.cellIndex, record]
      }
    }
  }
}
