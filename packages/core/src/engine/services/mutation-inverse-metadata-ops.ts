import type { EngineOp } from '@bilig/workbook-domain'
import { cloneCellStyleRecord } from '../../engine-style-utils.js'
import { restoreFormatRangeOps, restoreStyleRangeOps } from '../../engine-range-format-ops.js'
import type { WorkbookStore } from '../../workbook-store.js'
import { rangesIntersect } from '../../workbook-merge-records.js'

export function buildMutationMetadataInverseOps(workbook: WorkbookStore, op: EngineOp): EngineOp[] | undefined {
  switch (op.kind) {
    case 'upsertWorkbook':
      return [{ kind: 'upsertWorkbook', name: workbook.workbookName }]
    case 'setWorkbookMetadata': {
      const existing = workbook.getWorkbookProperty(op.key)
      return [{ kind: 'setWorkbookMetadata', key: op.key, value: existing?.value ?? null }]
    }
    case 'setCalculationSettings':
      return [{ kind: 'setCalculationSettings', settings: workbook.getCalculationSettings() }]
    case 'setVolatileContext':
      return [{ kind: 'setVolatileContext', context: workbook.getVolatileContext() }]
    case 'upsertSheet': {
      const existing = workbook.getSheet(op.name)
      if (!existing) {
        return [{ kind: 'deleteSheet', name: op.name }]
      }
      return [{ kind: 'upsertSheet', name: existing.name, order: existing.order }]
    }
    case 'renameSheet': {
      const existing = workbook.getSheet(op.newName)
      if (!existing) {
        return []
      }
      return [{ kind: 'renameSheet', oldName: op.newName, newName: op.oldName }]
    }
    case 'insertRows':
      return [{ kind: 'deleteRows', sheetName: op.sheetName, start: op.start, count: op.count }]
    case 'moveRows':
      return [{ kind: 'moveRows', sheetName: op.sheetName, start: op.target, count: op.count, target: op.start }]
    case 'insertColumns':
      return [{ kind: 'deleteColumns', sheetName: op.sheetName, start: op.start, count: op.count }]
    case 'moveColumns':
      return [{ kind: 'moveColumns', sheetName: op.sheetName, start: op.target, count: op.count, target: op.start }]
    case 'updateRowMetadata': {
      const existing = workbook.getRowMetadata(op.sheetName, op.start, op.count)
      return [
        {
          kind: 'updateRowMetadata',
          sheetName: op.sheetName,
          start: op.start,
          count: op.count,
          size: existing?.size ?? null,
          hidden: existing?.hidden ?? null,
        },
      ]
    }
    case 'updateColumnMetadata': {
      const existing = workbook.getColumnMetadata(op.sheetName, op.start, op.count)
      return [
        {
          kind: 'updateColumnMetadata',
          sheetName: op.sheetName,
          start: op.start,
          count: op.count,
          size: existing?.size ?? null,
          hidden: existing?.hidden ?? null,
        },
      ]
    }
    case 'setFreezePane': {
      const existing = workbook.getFreezePane(op.sheetName)
      if (!existing) {
        return [{ kind: 'clearFreezePane', sheetName: op.sheetName }]
      }
      return [{ kind: 'setFreezePane', sheetName: op.sheetName, rows: existing.rows, cols: existing.cols }]
    }
    case 'clearFreezePane': {
      const existing = workbook.getFreezePane(op.sheetName)
      if (!existing) {
        return []
      }
      return [{ kind: 'setFreezePane', sheetName: op.sheetName, rows: existing.rows, cols: existing.cols }]
    }
    case 'mergeCells': {
      const existing = workbook.listMergeRanges(op.range.sheetName).filter((record) => rangesIntersect(record, op.range))
      return [
        { kind: 'unmergeCells', range: { ...op.range } },
        ...existing.map((record) => ({
          kind: 'mergeCells' as const,
          range: { ...record },
        })),
      ]
    }
    case 'unmergeCells':
      return workbook
        .listMergeRanges(op.range.sheetName)
        .filter((record) => rangesIntersect(record, op.range))
        .map((record) => ({
          kind: 'mergeCells' as const,
          range: { ...record },
        }))
    case 'setFilter': {
      const existing = workbook.getFilter(op.sheetName, op.range)
      if (!existing) {
        return [{ kind: 'clearFilter', sheetName: op.sheetName, range: structuredClone(op.range) }]
      }
      return [{ kind: 'setFilter', sheetName: op.sheetName, range: structuredClone(existing.range) }]
    }
    case 'clearFilter': {
      const existing = workbook.getFilter(op.sheetName, op.range)
      if (!existing) {
        return []
      }
      return [{ kind: 'setFilter', sheetName: op.sheetName, range: structuredClone(existing.range) }]
    }
    case 'setSort': {
      const existing = workbook.getSort(op.sheetName, op.range)
      if (!existing) {
        return [{ kind: 'clearSort', sheetName: op.sheetName, range: { ...op.range } }]
      }
      return [
        {
          kind: 'setSort',
          sheetName: op.sheetName,
          range: { ...existing.range },
          keys: existing.keys.map((key) => Object.assign({}, key)),
        },
      ]
    }
    case 'clearSort': {
      const existing = workbook.getSort(op.sheetName, op.range)
      if (!existing) {
        return []
      }
      return [
        {
          kind: 'setSort',
          sheetName: op.sheetName,
          range: { ...existing.range },
          keys: existing.keys.map((key) => Object.assign({}, key)),
        },
      ]
    }
    case 'setDataValidation': {
      const existing = workbook.getDataValidation(op.validation.range.sheetName, op.validation.range)
      if (!existing) {
        return [{ kind: 'clearDataValidation', sheetName: op.validation.range.sheetName, range: { ...op.validation.range } }]
      }
      return [{ kind: 'setDataValidation', validation: structuredClone(existing) }]
    }
    case 'clearDataValidation': {
      const existing = workbook.getDataValidation(op.sheetName, op.range)
      if (!existing) {
        return []
      }
      return [{ kind: 'setDataValidation', validation: structuredClone(existing) }]
    }
    case 'setSheetProtection': {
      const existing = workbook.getSheetProtection(op.protection.sheetName)
      if (!existing) {
        return [{ kind: 'clearSheetProtection', sheetName: op.protection.sheetName }]
      }
      return [{ kind: 'setSheetProtection', protection: structuredClone(existing) }]
    }
    case 'clearSheetProtection': {
      const existing = workbook.getSheetProtection(op.sheetName)
      if (!existing) {
        return []
      }
      return [{ kind: 'setSheetProtection', protection: structuredClone(existing) }]
    }
    case 'upsertConditionalFormat': {
      const existing = workbook.getConditionalFormat(op.format.id)
      if (!existing) {
        return [{ kind: 'deleteConditionalFormat', id: op.format.id, sheetName: op.format.range.sheetName }]
      }
      return [{ kind: 'upsertConditionalFormat', format: structuredClone(existing) }]
    }
    case 'deleteConditionalFormat': {
      const existing = workbook.getConditionalFormat(op.id)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertConditionalFormat', format: structuredClone(existing) }]
    }
    case 'upsertRangeProtection': {
      const existing = workbook.getRangeProtection(op.protection.id)
      if (!existing) {
        return [{ kind: 'deleteRangeProtection', id: op.protection.id, sheetName: op.protection.range.sheetName }]
      }
      return [{ kind: 'upsertRangeProtection', protection: structuredClone(existing) }]
    }
    case 'deleteRangeProtection': {
      const existing = workbook.getRangeProtection(op.id)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertRangeProtection', protection: structuredClone(existing) }]
    }
    case 'upsertCommentThread': {
      const existing = workbook.getCommentThread(op.thread.sheetName, op.thread.address)
      if (!existing) {
        return [{ kind: 'deleteCommentThread', sheetName: op.thread.sheetName, address: op.thread.address }]
      }
      return [{ kind: 'upsertCommentThread', thread: structuredClone(existing) }]
    }
    case 'deleteCommentThread': {
      const existing = workbook.getCommentThread(op.sheetName, op.address)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertCommentThread', thread: structuredClone(existing) }]
    }
    case 'upsertNote': {
      const existing = workbook.getNote(op.note.sheetName, op.note.address)
      if (!existing) {
        return [{ kind: 'deleteNote', sheetName: op.note.sheetName, address: op.note.address }]
      }
      return [{ kind: 'upsertNote', note: structuredClone(existing) }]
    }
    case 'deleteNote': {
      const existing = workbook.getNote(op.sheetName, op.address)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertNote', note: structuredClone(existing) }]
    }
    case 'setCellFormat': {
      const cellIndex = workbook.getCellIndex(op.sheetName, op.address)
      return [
        {
          kind: 'setCellFormat',
          sheetName: op.sheetName,
          address: op.address,
          format: cellIndex === undefined ? null : (workbook.getCellFormat(cellIndex) ?? null),
        },
      ]
    }
    case 'upsertCellStyle': {
      const existing = workbook.getCellStyle(op.style.id)
      if (!existing || existing.id !== op.style.id) {
        return []
      }
      return [{ kind: 'upsertCellStyle', style: cloneCellStyleRecord(existing) }]
    }
    case 'upsertCellNumberFormat': {
      const existing = workbook.getCellNumberFormat(op.format.id)
      if (!existing || existing.id !== op.format.id) {
        return []
      }
      return [{ kind: 'upsertCellNumberFormat', format: { ...existing } }]
    }
    case 'setStyleRange':
      return restoreStyleRangeOps(workbook, op.range)
    case 'setFormatRange':
      return restoreFormatRangeOps(workbook, op.range)
    case 'upsertDefinedName': {
      const existing = workbook.getDefinedName(op.name)
      if (!existing) {
        return [{ kind: 'deleteDefinedName', name: op.name }]
      }
      return [{ kind: 'upsertDefinedName', name: existing.name, value: existing.value }]
    }
    case 'deleteDefinedName': {
      const existing = workbook.getDefinedName(op.name)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertDefinedName', name: existing.name, value: existing.value }]
    }
    case 'upsertTable': {
      const existing = workbook.getTable(op.table.name)
      if (!existing) {
        return [{ kind: 'deleteTable', name: op.table.name }]
      }
      return [
        {
          kind: 'upsertTable',
          table: {
            name: existing.name,
            sheetName: existing.sheetName,
            startAddress: existing.startAddress,
            endAddress: existing.endAddress,
            columnNames: [...existing.columnNames],
            headerRow: existing.headerRow,
            totalsRow: existing.totalsRow,
          },
        },
      ]
    }
    case 'deleteTable': {
      const existing = workbook.getTable(op.name)
      if (!existing) {
        return []
      }
      return [
        {
          kind: 'upsertTable',
          table: {
            name: existing.name,
            sheetName: existing.sheetName,
            startAddress: existing.startAddress,
            endAddress: existing.endAddress,
            columnNames: [...existing.columnNames],
            headerRow: existing.headerRow,
            totalsRow: existing.totalsRow,
          },
        },
      ]
    }
    case 'upsertSpillRange': {
      const existing = workbook.getSpill(op.sheetName, op.address)
      if (!existing) {
        return [{ kind: 'deleteSpillRange', sheetName: op.sheetName, address: op.address }]
      }
      return [
        { kind: 'upsertSpillRange', sheetName: existing.sheetName, address: existing.address, rows: existing.rows, cols: existing.cols },
      ]
    }
    case 'deleteSpillRange': {
      const existing = workbook.getSpill(op.sheetName, op.address)
      if (!existing) {
        return []
      }
      return [
        { kind: 'upsertSpillRange', sheetName: existing.sheetName, address: existing.address, rows: existing.rows, cols: existing.cols },
      ]
    }
    case 'upsertPivotTable': {
      const existing = workbook.getPivot(op.sheetName, op.address)
      if (!existing) {
        return [{ kind: 'deletePivotTable', sheetName: op.sheetName, address: op.address }]
      }
      return [
        {
          kind: 'upsertPivotTable',
          name: existing.name,
          sheetName: existing.sheetName,
          address: existing.address,
          source: { ...existing.source },
          groupBy: [...existing.groupBy],
          values: existing.values.map((value) => Object.assign({}, value)),
          rows: existing.rows,
          cols: existing.cols,
        },
      ]
    }
    case 'deletePivotTable': {
      const existing = workbook.getPivot(op.sheetName, op.address)
      if (!existing) {
        return []
      }
      return [
        {
          kind: 'upsertPivotTable',
          name: existing.name,
          sheetName: existing.sheetName,
          address: existing.address,
          source: { ...existing.source },
          groupBy: [...existing.groupBy],
          values: existing.values.map((value) => Object.assign({}, value)),
          rows: existing.rows,
          cols: existing.cols,
        },
      ]
    }
    case 'upsertChart': {
      const existing = workbook.getChart(op.chart.id)
      if (!existing) {
        return [{ kind: 'deleteChart', id: op.chart.id }]
      }
      return [{ kind: 'upsertChart', chart: structuredClone(existing) }]
    }
    case 'deleteChart': {
      const existing = workbook.getChart(op.id)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertChart', chart: structuredClone(existing) }]
    }
    case 'upsertImage': {
      const existing = workbook.getImage(op.image.id)
      if (!existing) {
        return [{ kind: 'deleteImage', id: op.image.id }]
      }
      return [{ kind: 'upsertImage', image: structuredClone(existing) }]
    }
    case 'deleteImage': {
      const existing = workbook.getImage(op.id)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertImage', image: structuredClone(existing) }]
    }
    case 'upsertShape': {
      const existing = workbook.getShape(op.shape.id)
      if (!existing) {
        return [{ kind: 'deleteShape', id: op.shape.id }]
      }
      return [{ kind: 'upsertShape', shape: structuredClone(existing) }]
    }
    case 'deleteShape': {
      const existing = workbook.getShape(op.id)
      if (!existing) {
        return []
      }
      return [{ kind: 'upsertShape', shape: structuredClone(existing) }]
    }
    case 'deleteSheet':
    case 'deleteRows':
    case 'deleteColumns':
    case 'setCellValue':
    case 'setCellFormula':
    case 'clearCell':
      return undefined
    default: {
      const exhaustive: never = op
      return exhaustive
    }
  }
}
