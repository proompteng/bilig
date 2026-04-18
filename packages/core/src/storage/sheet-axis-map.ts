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

  list(axis: AxisKind): AxisEntrySnapshot[] {
    return this.get(axis).list()
  }

  snapshot(axis: AxisKind, start: number, count: number): AxisEntrySnapshot[] {
    return this.get(axis).snapshot(start, count)
  }

  replaceRange(axis: AxisKind, start: number, entries: readonly AxisEntrySnapshot[]): void {
    this.get(axis).replaceRange(start, entries)
  }

  splice(axis: AxisKind, start: number, deleteCount: number, entries: readonly AxisEntrySnapshot[]): AxisEntrySnapshot[] {
    return this.get(axis).splice(start, deleteCount, entries)
  }

  move(axis: AxisKind, start: number, count: number, target: number): void {
    this.get(axis).move(start, count, target)
  }
}
