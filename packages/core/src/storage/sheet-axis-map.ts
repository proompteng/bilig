import type { AxisEntrySnapshot, AxisKind } from './axis-map.js'
import { AxisMap } from './axis-map.js'

export class SheetAxisMap {
  readonly rows: AxisMap
  readonly columns: AxisMap

  constructor() {
    this.rows = new AxisMap()
    this.columns = new AxisMap()
  }

  get(axis: AxisKind): AxisMap {
    return axis === 'row' ? this.rows : this.columns
  }

  getId(axis: AxisKind, index: number): string | undefined {
    return this.get(axis).getId(index)
  }

  setId(axis: AxisKind, index: number, id: string): void {
    this.get(axis).setId(index, id)
  }

  ensureId(axis: AxisKind, index: number, createId: () => string): string {
    return this.get(axis).ensureId(index, createId)
  }

  indexOf(axis: AxisKind, id: string): number {
    return this.get(axis).indexOf(id)
  }

  list(axis: AxisKind): AxisEntrySnapshot[] {
    return this.get(axis).list()
  }

  snapshot(axis: AxisKind, start: number, count: number): AxisEntrySnapshot[] {
    return this.get(axis).snapshot(start, count)
  }

  replaceRange(axis: AxisKind, start: number, entries: readonly AxisEntrySnapshot[]): void {
    this.get(axis).replaceRange(start, entries)
  }

  splice(axis: AxisKind, start: number, deleteCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[]
  splice(
    axis: AxisKind,
    start: number,
    deleteCount: number,
    insertCount: number,
    entries: readonly AxisEntrySnapshot[],
  ): AxisEntrySnapshot[]
  splice(
    axis: AxisKind,
    start: number,
    deleteCount: number,
    insertCountOrEntries: number | readonly AxisEntrySnapshot[],
    maybeEntries?: readonly AxisEntrySnapshot[],
  ): AxisEntrySnapshot[] {
    if (typeof insertCountOrEntries === 'number') {
      return this.get(axis).splice(start, deleteCount, insertCountOrEntries, maybeEntries ?? [])
    }
    return this.get(axis).splice(start, deleteCount, insertCountOrEntries)
  }

  move(axis: AxisKind, start: number, count: number, target: number): void {
    this.get(axis).move(start, count, target)
  }
}
