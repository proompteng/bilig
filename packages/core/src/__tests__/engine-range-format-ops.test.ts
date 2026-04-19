import { describe, expect, it } from 'vitest'
import { createCellNumberFormatRecord } from '@bilig/protocol'
import {
  buildFormatPatchOps,
  buildStyleClearOps,
  buildStylePatchOps,
  restoreFormatRangeOps,
  restoreStyleRangeOps,
} from '../engine-range-format-ops.js'
import { WorkbookStore } from '../workbook-store.js'

describe('engine range format ops', () => {
  it('splits style tiles and upserts patched styles across overlapping ranges', () => {
    const workbook = new WorkbookStore('style-ops')
    workbook.createSheet('Sheet1')
    workbook.upsertCellStyle({ id: 'style-a', font: { bold: true } })
    workbook.upsertCellStyle({ id: 'style-b', alignment: { horizontal: 'center' } })
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, 'style-a')
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, 'style-b')

    const ops = buildStylePatchOps(
      workbook,
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C2' },
      { fill: { backgroundColor: '#ffeedd' } },
    )

    expect(ops).toHaveLength(4)
    expect(ops).toMatchObject([
      {
        kind: 'upsertCellStyle',
        style: {
          font: { bold: true },
          fill: { backgroundColor: '#ffeedd' },
        },
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      },
      {
        kind: 'upsertCellStyle',
        style: {
          alignment: { horizontal: 'center' },
          fill: { backgroundColor: '#ffeedd' },
        },
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' },
      },
    ])
  })

  it('clears selected style fields and patches number formats across tiled ranges', () => {
    const workbook = new WorkbookStore('format-ops')
    workbook.createSheet('Sheet1')
    workbook.upsertCellStyle({
      id: 'style-c',
      fill: { backgroundColor: '#ffffff' },
      font: { family: 'Inter', bold: true },
    })
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, 'style-c')
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-money', '$0.00'))
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, 'format-money')

    const clearedStyleOps = buildStyleClearOps(workbook, { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, ['fontBold'])
    expect(clearedStyleOps).toHaveLength(2)
    expect(clearedStyleOps).toMatchObject([
      {
        kind: 'upsertCellStyle',
        style: {
          fill: { backgroundColor: '#ffffff' },
          font: { family: 'Inter' },
        },
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      },
    ])

    const formatOps = buildFormatPatchOps(workbook, { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, '0.00')
    expect(formatOps).toHaveLength(3)
    expect(formatOps).toMatchObject([
      {
        kind: 'upsertCellNumberFormat',
        format: { code: '0.00' },
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B2' },
      },
    ])
  })

  it('coalesces adjacent replayed style tiles without coalescing format tiles', () => {
    const workbook = new WorkbookStore('restore-range-ops')
    workbook.createSheet('Sheet1')
    workbook.upsertCellStyle({ id: 'style-fill', fill: { backgroundColor: '#dbeafe' } })
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-decimal', '0.00'))

    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A3' }, 'style-fill')
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'A4' }, 'style-fill')
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'B3' }, 'format-decimal')
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'B4' }, 'format-decimal')

    expect(restoreStyleRangeOps(workbook, { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4' })).toMatchObject([
      {
        kind: 'upsertCellStyle',
        style: { id: 'style-fill', fill: { backgroundColor: '#dbeafe' } },
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4' },
        styleId: 'style-fill',
      },
    ])

    expect(restoreFormatRangeOps(workbook, { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'B4' })).toMatchObject([
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format-decimal', code: '0.00' },
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'B3' },
        formatId: 'format-decimal',
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'B4', endAddress: 'B4' },
        formatId: 'format-decimal',
      },
    ])
  })

  it('restores overlapping style ranges with their original extents while clearing untouched cells to default', () => {
    const workbook = new WorkbookStore('restore-overlapping-style-extents')
    workbook.createSheet('Sheet1')
    workbook.upsertCellStyle({ id: 'style-fill', fill: { backgroundColor: '#dbeafe' } })

    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'C3' }, 'style-fill')
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' }, 'style-fill')

    expect(restoreStyleRangeOps(workbook, { sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D4' })).toMatchObject([
      {
        kind: 'upsertCellStyle',
        style: { id: 'style-fill', fill: { backgroundColor: '#dbeafe' } },
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' },
        styleId: 'style-fill',
      },
      {
        kind: 'setStyleRange',
        range: { sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D3' },
        styleId: 'style-0',
      },
    ])
  })

  it('restores overlapping format ranges with their original extents while clearing untouched cells to default', () => {
    const workbook = new WorkbookStore('restore-overlapping-format-extents')
    workbook.createSheet('Sheet1')
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-decimal', '0.00'))

    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'C3' }, 'format-decimal')
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' }, 'format-decimal')

    expect(restoreFormatRangeOps(workbook, { sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D4' })).toMatchObject([
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format-decimal', code: '0.00' },
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'C4', endAddress: 'D4' },
        formatId: 'format-decimal',
      },
      {
        kind: 'setFormatRange',
        range: { sheetName: 'Sheet1', startAddress: 'D3', endAddress: 'D3' },
        formatId: 'format-0',
      },
    ])
  })
})
