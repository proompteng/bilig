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
      const index = Math.max(start, entry.index)
      this.entries[index] = entry.id
    }
  }

  splice(start: number, deleteCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[] {
    const removed = this.entries.splice(start, deleteCount, ...Array.from({ length: entries.length }, (_, index) => entries[index]?.id))
    return removed.flatMap((id, index) => (id === undefined ? [] : [createAxisEntrySnapshot(id, start + index)]))
  }

  move(start: number, count: number, target: number): void {
    if (count <= 0 || start === target) {
      return
    }
    const moved = this.entries.splice(start, count)
    this.entries.splice(target, 0, ...moved)
  }
}
