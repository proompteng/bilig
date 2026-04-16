import { describe, expect, it } from 'vitest'
import { createCellNumberFormatRecord } from '@bilig/protocol'
import { exportSheetMetadata, sheetMetadataToOps } from '../engine-snapshot-utils.js'
import { WorkbookStore } from '../workbook-store.js'

describe('engine snapshot utils', () => {
  it('returns no sheet metadata when a sheet has no persisted metadata', () => {
    const workbook = new WorkbookStore('snapshot-empty')
    workbook.createSheet('Sheet1')

    expect(exportSheetMetadata(workbook, 'Sheet1')).toBeUndefined()
    expect(sheetMetadataToOps(workbook, 'Sheet1')).toEqual([])
  })

  it('serializes and restores workbook sheet metadata with cloned records', () => {
    const workbook = new WorkbookStore('snapshot-spec')
    workbook.createSheet('Sheet1')
    workbook.insertRows('Sheet1', 0, 1, [{ id: 'row-1', index: 0, size: 24, hidden: true }])
    workbook.insertColumns('Sheet1', 1, 1, [{ id: 'column-2', index: 1, size: 140, hidden: false }])
    workbook.setRowMetadata('Sheet1', 2, 2, 30, true)
    workbook.setColumnMetadata('Sheet1', 3, 1, 160, false)
    workbook.upsertCellStyle({ id: 'style-bold', font: { bold: true } })
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, 'style-bold')
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-decimal', '0.00'))
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C3' }, 'format-decimal')
    workbook.setFreezePane('Sheet1', 1, 2)
    workbook.setFilter('Sheet1', {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D8',
    })
    workbook.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D8' }, [{ keyAddress: 'B2', direction: 'desc' }])

    const metadata = exportSheetMetadata(workbook, 'Sheet1')
    expect(metadata).toEqual({
      rows: [
        { id: 'row-1', index: 0, size: 24, hidden: true },
        { id: 'row-1', index: 2, size: 30, hidden: true },
        { id: 'row-2', index: 3, size: 30, hidden: true },
      ],
      columns: [
        { id: 'column-2', index: 1, size: 140, hidden: false },
        { id: 'column-1', index: 3, size: 160, hidden: false },
      ],
      rowMetadata: [
        { start: 0, count: 1, size: 24, hidden: true },
        { start: 2, count: 2, size: 30, hidden: true },
      ],
      columnMetadata: [
        { start: 1, count: 1, size: 140, hidden: false },
        { start: 3, count: 1, size: 160, hidden: false },
      ],
      styleRanges: [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
          styleId: 'style-bold',
        },
      ],
      formatRanges: [
        {
          range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C3' },
          formatId: 'format-decimal',
        },
      ],
      freezePane: { rows: 1, cols: 2 },
      filters: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D8' }],
      sorts: [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D8' },
          keys: [{ keyAddress: 'B2', direction: 'desc' }],
        },
      ],
    })

    if (!metadata) {
      throw new Error('Expected sheet metadata')
    }
    metadata.styleRanges![0].range.startAddress = 'Z9'
    metadata.sorts![0].keys[0].direction = 'asc'

    expect(workbook.listStyleRanges('Sheet1')[0]?.range.startAddress).toBe('A1')
    expect(workbook.listSorts('Sheet1')[0]?.keys[0]?.direction).toBe('desc')

    const ops = sheetMetadataToOps(workbook, 'Sheet1')
    expect(ops).toHaveLength(14)
    expect(ops).toMatchObject([
      {
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
        entries: [{ id: 'row-1', index: 0, size: 24, hidden: true }],
      },
      {
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 2,
        count: 1,
        entries: [{ id: 'row-1', index: 2, size: 30, hidden: true }],
      },
      {
        kind: 'insertRows',
        sheetName: 'Sheet1',
        start: 3,
        count: 1,
        entries: [{ id: 'row-2', index: 3, size: 30, hidden: true }],
      },
      {
        kind: 'insertColumns',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
        entries: [{ id: 'column-2', index: 1, size: 140, hidden: false }],
      },
      {
        kind: 'insertColumns',
        sheetName: 'Sheet1',
        start: 3,
        count: 1,
        entries: [{ id: 'column-1', index: 3, size: 160, hidden: false }],
      },
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        start: 0,
        count: 1,
        size: 24,
        hidden: true,
      },
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        start: 2,
        count: 2,
        size: 30,
        hidden: true,
      },
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        start: 1,
        count: 1,
        size: 140,
        hidden: false,
      },
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        start: 3,
        count: 1,
        size: 160,
        hidden: false,
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        styleId: 'style-bold',
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C3' },
        formatId: 'format-decimal',
      },
      {
        kind: 'setFreezePane',
        sheetName: 'Sheet1',
        rows: 1,
        cols: 2,
      },
      {
        kind: 'setFilter',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D8' },
      },
      {
        kind: 'setSort',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D8' },
        keys: [{ keyAddress: 'B2', direction: 'desc' }],
      },
    ])
  })
})
