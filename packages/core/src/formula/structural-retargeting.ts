import type { FormulaInstanceSnapshot } from './formula-instance-table.js'

export function retargetFormulaInstance(
  record: FormulaInstanceSnapshot,
  next: {
    readonly sheetName?: string
    readonly row?: number
    readonly col?: number
    readonly source?: string
    readonly templateId?: number
  },
): FormulaInstanceSnapshot {
  return {
    cellIndex: record.cellIndex,
    sheetName: next.sheetName ?? record.sheetName,
    row: next.row ?? record.row,
    col: next.col ?? record.col,
    source: next.source ?? record.source,
    ...(next.templateId !== undefined
      ? { templateId: next.templateId }
      : record.templateId !== undefined
        ? { templateId: record.templateId }
        : {}),
  }
}
