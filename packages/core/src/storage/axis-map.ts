export type AxisKind = 'row' | 'column'

export interface AxisEntrySnapshot {
  readonly id: string
  readonly index: number
}

function createAxisEntrySnapshot(id: string, index: number): AxisEntrySnapshot {
  return { id, index }
}

export class AxisMap {
  private readonly entries: Array<string | undefined> = []
  private readonly idToIndex = new Map<string, number>()

  get(index: number): string | undefined {
    return this.entries[index]
  }

  getId(index: number): string | undefined {
    return this.get(index)
  }

  set(index: number, id: string): void {
    const previous = this.entries[index]
    if (previous !== undefined) {
      this.idToIndex.delete(previous)
    }
    this.entries[index] = id
    this.idToIndex.set(id, index)
  }

  setId(index: number, id: string): void {
    this.set(index, id)
  }

  ensure(index: number, createId: () => string): string {
    const existing = this.entries[index]
    if (existing !== undefined) {
      return existing
    }
    if (index >= this.entries.length) {
      this.entries.length = index + 1
    }
    const id = createId()
    this.entries[index] = id
    this.idToIndex.set(id, index)
    return id
  }

  ensureId(index: number, createId: () => string): string {
    return this.ensure(index, createId)
  }

  indexOf(id: string): number {
    return this.idToIndex.get(id) ?? -1
  }

  get length(): number {
    return this.entries.length
  }

  list(): AxisEntrySnapshot[] {
    const snapshots: AxisEntrySnapshot[] = []
    for (let index = 0; index < this.entries.length; index += 1) {
      const id = this.entries[index]
      if (id === undefined) {
        continue
      }
      snapshots.push(createAxisEntrySnapshot(id, index))
    }
    return snapshots
  }

  snapshot(start: number, count: number): AxisEntrySnapshot[] {
    if (count <= 0) {
      return []
    }
    const snapshots: AxisEntrySnapshot[] = []
    for (let offset = 0; offset < count; offset += 1) {
      const index = start + offset
      const id = this.entries[index]
      if (id === undefined) {
        continue
      }
      snapshots.push(createAxisEntrySnapshot(id, index))
    }
    return snapshots
  }

  replaceRange(start: number, entries: readonly AxisEntrySnapshot[]): void {
    for (const entry of entries) {
      if (entry.index < start) {
        continue
      }
      const previous = this.entries[entry.index]
      if (previous !== undefined) {
        this.idToIndex.delete(previous)
      }
      this.entries[entry.index] = entry.id
      this.idToIndex.set(entry.id, entry.index)
    }
  }

  splice(start: number, deleteCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[]
  splice(start: number, deleteCount: number, insertCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[]
  splice(
    start: number,
    deleteCount: number,
    insertCountOrEntries: number | readonly AxisEntrySnapshot[],
    maybeEntries?: readonly AxisEntrySnapshot[],
  ): AxisEntrySnapshot[] {
    const entries = typeof insertCountOrEntries === 'number' ? (maybeEntries ?? []) : insertCountOrEntries
    const explicitInsertCount = typeof insertCountOrEntries === 'number' ? insertCountOrEntries : entries.length
    if (this.entries.length < start) {
      this.entries.length = start
    }
    const insertLength = Math.max(
      explicitInsertCount,
      entries.reduce((max, entry) => Math.max(max, entry.index - start + 1), 0),
    )
    const inserted: Array<string | undefined> = Array.from({ length: insertLength }, () => undefined)
    for (const entry of entries) {
      const offset = entry.index - start
      if (offset < 0 || offset >= insertLength) {
        continue
      }
      inserted[offset] = entry.id
    }
    const removed = this.entries.splice(start, deleteCount, ...inserted)
    this.rebuildIndex()
    return removed.flatMap((id, index) => (id === undefined ? [] : [createAxisEntrySnapshot(id, start + index)]))
  }

  move(start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return
    }
    const moved = this.entries.splice(start, count)
    this.entries.splice(target, 0, ...moved)
    this.rebuildIndex()
  }

  private rebuildIndex(): void {
    this.idToIndex.clear()
    for (let index = 0; index < this.entries.length; index += 1) {
      const id = this.entries[index]
      if (id !== undefined) {
        this.idToIndex.set(id, index)
      }
    }
  }
}
