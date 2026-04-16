import { describe, expect, it } from 'vitest'
import { createCellNumberFormatRecord } from '@bilig/protocol'
import { buildFormatPatchOps, buildStyleClearOps, buildStylePatchOps } from '../engine-range-format-ops.js'
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
})
