import { describe, expect, it } from 'vitest'
import { WorkbookStore } from '../workbook-store.js'
import {
  buildDeleteChartOps,
  buildDeleteTableOps,
  buildSetChartOps,
  buildSetTableOps,
} from '../engine/engine-workbook-object-metadata-ops.js'

describe('engine workbook object metadata ops', () => {
  it('builds idempotent table upserts and existing-table deletes', () => {
    const workbook = new WorkbookStore()
    const table = {
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    }

    const setOps = buildSetTableOps(workbook, table)
    expect(setOps).toEqual([{ kind: 'upsertTable', table }])
    workbook.setTable(table)

    expect(buildSetTableOps(workbook, { ...table, columnNames: [...table.columnNames] })).toEqual([])
    expect(buildDeleteTableOps(workbook, 'Missing')).toBeNull()
    expect(buildDeleteTableOps(workbook, 'Sales')).toEqual([{ kind: 'deleteTable', name: 'Sales' }])
  })

  it('clones chart upsert payloads before handing them to transactions', () => {
    const workbook = new WorkbookStore()
    const chart = {
      id: 'RevenueChart',
      sheetName: 'Sheet1',
      address: 'D2',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
      chartType: 'column' as const,
      rows: 8,
      cols: 6,
      title: 'Revenue',
    }

    const setOps = buildSetChartOps(workbook, chart)
    chart.source.startAddress = 'Z9'

    expect(setOps).toEqual([
      {
        kind: 'upsertChart',
        chart: {
          id: 'RevenueChart',
          sheetName: 'Sheet1',
          address: 'D2',
          source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' },
          chartType: 'column',
          rows: 8,
          cols: 6,
          title: 'Revenue',
        },
      },
    ])

    const chartOp = setOps[0]
    if (chartOp?.kind !== 'upsertChart') {
      throw new Error('Expected chart upsert op')
    }
    workbook.setChart(chartOp.chart)
    expect(buildSetChartOps(workbook, chartOp.chart)).toEqual([])
    expect(buildDeleteChartOps(workbook, 'Missing')).toBeNull()
    expect(buildDeleteChartOps(workbook, 'RevenueChart')).toEqual([{ kind: 'deleteChart', id: 'RevenueChart' }])
  })
})
