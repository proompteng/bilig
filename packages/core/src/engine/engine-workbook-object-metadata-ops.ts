import type { WorkbookChartSnapshot, WorkbookImageSnapshot, WorkbookShapeSnapshot } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import type { WorkbookStore, WorkbookTableRecord } from '../workbook-store.js'
import {
  cloneEngineTableRecord,
  workbookChartsEqual,
  workbookImagesEqual,
  workbookShapesEqual,
  workbookTablesEqual,
} from './engine-workbook-object-helpers.js'

export function buildSetTableOps(workbook: WorkbookStore, table: WorkbookTableRecord): EngineOp[] {
  const existing = workbook.getTable(table.name)
  if (workbookTablesEqual(existing, table)) {
    return []
  }
  return [
    {
      kind: 'upsertTable',
      table: cloneEngineTableRecord(table),
    },
  ]
}

export function buildDeleteTableOps(workbook: WorkbookStore, name: string): EngineOp[] | null {
  if (!workbook.getTable(name)) {
    return null
  }
  return [{ kind: 'deleteTable', name }]
}

export function buildSetChartOps(workbook: WorkbookStore, chart: WorkbookChartSnapshot): EngineOp[] {
  const existing = workbook.getChart(chart.id)
  if (workbookChartsEqual(existing, chart)) {
    return []
  }
  return [{ kind: 'upsertChart', chart: structuredClone(chart) }]
}

export function buildDeleteChartOps(workbook: WorkbookStore, id: string): EngineOp[] | null {
  if (!workbook.getChart(id)) {
    return null
  }
  return [{ kind: 'deleteChart', id }]
}

export function buildSetImageOps(workbook: WorkbookStore, image: WorkbookImageSnapshot): EngineOp[] {
  const existing = workbook.getImage(image.id)
  if (workbookImagesEqual(existing, image)) {
    return []
  }
  return [{ kind: 'upsertImage', image: structuredClone(image) }]
}

export function buildDeleteImageOps(workbook: WorkbookStore, id: string): EngineOp[] | null {
  if (!workbook.getImage(id)) {
    return null
  }
  return [{ kind: 'deleteImage', id }]
}

export function buildSetShapeOps(workbook: WorkbookStore, shape: WorkbookShapeSnapshot): EngineOp[] {
  const existing = workbook.getShape(shape.id)
  if (workbookShapesEqual(existing, shape)) {
    return []
  }
  return [{ kind: 'upsertShape', shape: structuredClone(shape) }]
}

export function buildDeleteShapeOps(workbook: WorkbookStore, id: string): EngineOp[] | null {
  if (!workbook.getShape(id)) {
    return null
  }
  return [{ kind: 'deleteShape', id }]
}
