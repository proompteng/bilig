export type AxisId = string

export interface AxisPieceSnapshot {
  readonly id: AxisId
  readonly index: number
}

export class AxisPieceMap {
  private readonly entries: Array<AxisId | undefined> = []

  get length(): number {
    return this.entries.length
  }

  getId(index: number): AxisId | undefined {
    return this.entries[index]
  }

  idAt(index: number): AxisId | undefined {
    return this.getId(index)
  }

  setId(index: number, id: AxisId): void {
    this.assertCanUseIds([{ id, index }], new Set([this.entries[index]].filter((entry): entry is AxisId => entry !== undefined)))
    if (index >= this.entries.length) {
      this.entries.length = index + 1
    }
    this.entries[index] = id
  }

  ensureId(index: number, createId: () => AxisId): AxisId {
    const existing = this.entries[index]
    if (existing !== undefined) {
      return existing
    }
    const id = createId()
    this.setId(index, id)
    return id
  }

  indexOf(id: AxisId): number {
    return this.entries.indexOf(id)
  }

  indexOfId(id: AxisId): number {
    return this.indexOf(id)
  }

  list(): AxisPieceSnapshot[] {
    return this.entries.flatMap((id, index) => (id === undefined ? [] : [{ id, index }]))
  }

  snapshot(start: number, count: number): AxisPieceSnapshot[] {
    if (count <= 0) {
      return []
    }
    const snapshots: AxisPieceSnapshot[] = []
    for (let offset = 0; offset < count; offset += 1) {
      const index = start + offset
      const id = this.entries[index]
      if (id !== undefined) {
        snapshots.push({ id, index })
      }
    }
    return snapshots
  }

  replace(ids: readonly AxisId[]): void {
    this.assertUniqueIds(ids)
    this.entries.splice(0, this.entries.length, ...ids)
  }

  replaceRange(start: number, entries: readonly AxisPieceSnapshot[]): void {
    this.assertCanUseIds(entries, new Set(entries.map((entry) => this.entries[entry.index]).filter((id): id is AxisId => id !== undefined)))
    for (const entry of entries) {
      if (entry.index < start) {
        continue
      }
      if (entry.index >= this.entries.length) {
        this.entries.length = entry.index + 1
      }
      this.entries[entry.index] = entry.id
    }
  }

  splice(start: number, deleteCount: number, insertedIds: readonly AxisId[]): AxisPieceSnapshot[]
  splice(start: number, deleteCount: number, insertCount: number, entries?: readonly AxisPieceSnapshot[]): AxisPieceSnapshot[]
  splice(
    start: number,
    deleteCount: number,
    insertCountOrInsertedIds: number | readonly AxisId[],
    maybeEntries: readonly AxisPieceSnapshot[] = [],
  ): AxisPieceSnapshot[] {
    const insertCount = typeof insertCountOrInsertedIds === 'number' ? insertCountOrInsertedIds : insertCountOrInsertedIds.length
    const entries =
      typeof insertCountOrInsertedIds === 'number'
        ? maybeEntries
        : insertCountOrInsertedIds.map((id, offset): AxisPieceSnapshot => ({ id, index: start + offset }))
    const removedIds = new Set(this.entries.slice(start, start + deleteCount).filter((id): id is AxisId => id !== undefined))
    this.assertCanUseIds(entries, removedIds)
    if (this.entries.length < start) {
      this.entries.length = start
    }
    const inserted: Array<AxisId | undefined> = Array.from({ length: insertCount }, () => undefined)
    for (const entry of entries) {
      const offset = entry.index - start
      if (offset < 0 || offset >= insertCount) {
        continue
      }
      inserted[offset] = entry.id
    }
    const removed = this.entries.splice(start, deleteCount, ...inserted)
    return removed.flatMap((id, offset) => (id === undefined ? [] : [{ id, index: start + offset }]))
  }

  delete(start: number, count: number): AxisPieceSnapshot[] {
    return this.splice(start, count, 0)
  }

  move(start: number, count: number, target: number): AxisPieceSnapshot[] {
    if (count <= 0 || start === target) {
      return []
    }
    const moved = this.entries.splice(start, count)
    this.entries.splice(target, 0, ...moved)
    return moved.flatMap((id, offset) => (id === undefined ? [] : [{ id, index: start + offset }]))
  }

  private assertUniqueIds(ids: readonly AxisId[]): void {
    if (new Set(ids).size !== ids.length) {
      throw new Error('Axis ids must be unique')
    }
  }

  private assertCanUseIds(entries: readonly AxisPieceSnapshot[], reusableIds: ReadonlySet<AxisId>): void {
    this.assertUniqueIds(entries.map((entry) => entry.id))
    for (const entry of entries) {
      if (!reusableIds.has(entry.id) && this.entries.includes(entry.id)) {
        throw new Error(`Axis id already exists: ${entry.id}`)
      }
    }
  }
}
