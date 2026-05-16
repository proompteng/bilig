import type { WorkbookAxisEntrySnapshot } from '@bilig/protocol'
import type { WorkbookAxisEntryStore } from './workbook-axis-entry-store.js'
import type { SheetRecord } from './workbook-sheet-record.js'

type WorkbookStructuralAxis = 'row' | 'column'

export interface WorkbookStructuralAxisOperationContext {
  readonly axisEntryStore: Pick<WorkbookAxisEntryStore, 'moveAxisEntries' | 'spliceAxisEntries'>
  readonly getOrCreateSheet: (sheetName: string) => SheetRecord
  readonly bumpSheetStructureVersion: (sheet: SheetRecord) => void
}

export class WorkbookStructuralAxisOperations {
  constructor(private readonly context: WorkbookStructuralAxisOperationContext) {}

  insert(
    axis: WorkbookStructuralAxis,
    sheetName: string,
    start: number,
    count: number,
    entries?: readonly WorkbookAxisEntrySnapshot[],
  ): void {
    const sheet = this.context.getOrCreateSheet(sheetName)
    this.context.axisEntryStore.spliceAxisEntries(sheet, axis, start, 0, count, entries)
    this.context.bumpSheetStructureVersion(sheet)
  }

  delete(axis: WorkbookStructuralAxis, sheetName: string, start: number, count: number): WorkbookAxisEntrySnapshot[] {
    const sheet = this.context.getOrCreateSheet(sheetName)
    const deleted = this.context.axisEntryStore.spliceAxisEntries(sheet, axis, start, count, 0)
    this.context.bumpSheetStructureVersion(sheet)
    return deleted
  }

  move(axis: WorkbookStructuralAxis, sheetName: string, start: number, count: number, target: number): void {
    const sheet = this.context.getOrCreateSheet(sheetName)
    this.context.axisEntryStore.moveAxisEntries(sheet, axis, start, count, target)
    this.context.bumpSheetStructureVersion(sheet)
  }
}
