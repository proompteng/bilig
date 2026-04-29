import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import { SpreadsheetEngine } from '../engine.js'

const protectedRange = {
  sheetName: 'Sheet1',
  startAddress: 'A1',
  endAddress: 'B2',
}

describe('operation-service metadata operations', () => {
  it('applies a mixed operation batch through the generic operation switch', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-applyops-coverage' })
    await engine.ready()

    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C4' }
    const ops: EngineOp[] = [
      { kind: 'upsertWorkbook', name: 'ApplyOps Book' },
      { kind: 'setWorkbookMetadata', key: 'owner', value: 'finance' },
      { kind: 'setCalculationSettings', settings: { mode: 'manual' } },
      { kind: 'setVolatileContext', context: { recalcEpoch: 9 } },
      { kind: 'upsertSheet', name: 'Sheet1', order: 0, id: 101 },
      { kind: 'upsertSheet', name: 'Temp', order: 1, id: 102 },
      { kind: 'renameSheet', oldName: 'Temp', newName: 'Archive' },
      { kind: 'upsertCellStyle', style: { id: 'style:accent', fill: { backgroundColor: '#ffeeaa' }, font: { bold: true } } },
      { kind: 'upsertCellNumberFormat', format: { id: 'fmt:currency', code: '$#,##0.00', kind: 'currency' } },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'Region' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 'Sales' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 'East' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B2', value: 10 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A3', value: 'West' },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B3', value: 20 },
      { kind: 'setCellFormula', sheetName: 'Sheet1', address: 'C2', formula: 'B2*2' },
      { kind: 'setCellFormat', sheetName: 'Sheet1', address: 'B2', format: '$0.00' },
      { kind: 'setStyleRange', range, styleId: 'style:accent' },
      { kind: 'setFormatRange', range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B3' }, formatId: 'fmt:currency' },
      { kind: 'updateRowMetadata', sheetName: 'Sheet1', start: 1, count: 1, size: 28, hidden: false },
      { kind: 'updateColumnMetadata', sheetName: 'Sheet1', start: 1, count: 1, size: 144, hidden: false },
      { kind: 'insertRows', sheetName: 'Sheet1', start: 4, count: 1 },
      { kind: 'deleteRows', sheetName: 'Sheet1', start: 4, count: 1 },
      { kind: 'insertColumns', sheetName: 'Sheet1', start: 3, count: 1 },
      { kind: 'deleteColumns', sheetName: 'Sheet1', start: 3, count: 1 },
      { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 1 },
      { kind: 'clearFreezePane', sheetName: 'Sheet1' },
      { kind: 'setFilter', sheetName: 'Sheet1', range },
      { kind: 'clearFilter', sheetName: 'Sheet1', range },
      { kind: 'setSort', sheetName: 'Sheet1', range, keys: [{ keyAddress: 'B1', direction: 'desc' }] },
      { kind: 'clearSort', sheetName: 'Sheet1', range },
      {
        kind: 'setDataValidation',
        validation: {
          id: 'validation:batch',
          range,
          rule: { kind: 'decimal', operator: 'greaterThan', values: [0] },
          allowBlank: true,
        },
      },
      { kind: 'clearDataValidation', sheetName: 'Sheet1', range },
      {
        kind: 'upsertConditionalFormat',
        format: {
          id: 'conditional:batch',
          range,
          rule: { kind: 'cellIs', operator: 'greaterThan', values: [10] },
          style: { font: { bold: true } },
        },
      },
      { kind: 'deleteConditionalFormat', id: 'conditional:batch', sheetName: 'Sheet1' },
      { kind: 'upsertRangeProtection', protection: { id: 'range-protect:batch', range } },
      { kind: 'deleteRangeProtection', id: 'range-protect:batch', sheetName: 'Sheet1' },
      { kind: 'setSheetProtection', protection: { sheetName: 'Sheet1' } },
      { kind: 'clearSheetProtection', sheetName: 'Sheet1' },
      {
        kind: 'upsertCommentThread',
        thread: {
          threadId: 'thread:batch',
          sheetName: 'Sheet1',
          address: 'A1',
          comments: [{ id: 'comment:batch', body: 'Looks good', authorUserId: 'u1', createdAtUnixMs: 1 }],
        },
      },
      { kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'A1' },
      { kind: 'upsertNote', note: { sheetName: 'Sheet1', address: 'A2', text: 'note' } },
      { kind: 'deleteNote', sheetName: 'Sheet1', address: 'A2' },
      { kind: 'upsertDefinedName', name: 'SalesTotal', value: { kind: 'formula', formula: 'SUM(Sheet1!B2:B3)' } },
      { kind: 'deleteDefinedName', name: 'SalesTotal' },
      {
        kind: 'upsertTable',
        table: {
          name: 'BatchTable',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'C3',
          columnNames: ['Region', 'Sales', 'Twice'],
          headerRow: true,
          totalsRow: false,
        },
      },
      { kind: 'deleteTable', name: 'BatchTable' },
      { kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 2 },
      { kind: 'deleteSpillRange', sheetName: 'Sheet1', address: 'E1' },
      {
        kind: 'upsertPivotTable',
        name: 'BatchPivot',
        sheetName: 'Sheet1',
        address: 'H1',
        source: range,
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Sales' }],
        rows: 4,
        cols: 3,
      },
      { kind: 'deletePivotTable', sheetName: 'Sheet1', address: 'H1' },
      {
        kind: 'upsertChart',
        chart: {
          id: 'chart:batch',
          sheetName: 'Sheet1',
          address: 'K1',
          source: range,
          chartType: 'bar',
          rows: 5,
          cols: 6,
        },
      },
      { kind: 'deleteChart', id: 'chart:batch' },
      {
        kind: 'upsertImage',
        image: { id: 'image:batch', sheetName: 'Sheet1', address: 'K8', sourceUrl: 'https://example.com/image.png', rows: 2, cols: 3 },
      },
      { kind: 'deleteImage', id: 'image:batch' },
      {
        kind: 'upsertShape',
        shape: { id: 'shape:batch', sheetName: 'Sheet1', address: 'K12', shapeType: 'rectangle', rows: 2, cols: 3 },
      },
      { kind: 'deleteShape', id: 'shape:batch' },
      { kind: 'clearCell', sheetName: 'Sheet1', address: 'C2' },
      { kind: 'deleteSheet', name: 'Archive' },
    ]

    const undoOps = engine.applyOps(ops, { captureUndo: true, source: 'local', potentialNewCells: 16 })

    expect(undoOps).not.toBeNull()
    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.name).toBe('ApplyOps Book')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Empty })
    expect(snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Sheet1'])
  })

  it('applies and removes workbook metadata objects through the public operation path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-metadata-coverage' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C3' }, [
      ['Region', 'Sales', 'Units'],
      ['East', 10, 1],
      ['West', 20, 2],
    ])

    engine.setFilter('Sheet1', protectedRange)
    engine.setSort('Sheet1', protectedRange, [{ keyAddress: 'B1', direction: 'desc' }])
    engine.setDataValidation({
      id: 'validation-1',
      range: protectedRange,
      rule: { kind: 'whole', operator: 'between', values: [1, 100] },
      allowBlank: false,
      showDropdown: true,
      promptTitle: 'Sales',
      promptMessage: 'Enter a sales amount',
      errorStyle: 'stop',
      errorTitle: 'Invalid',
      errorMessage: 'Sales must be in range',
    })
    engine.setConditionalFormat({
      id: 'format-1',
      range: protectedRange,
      rule: { kind: 'cellIs', operator: 'greaterThan', values: [15] },
      style: { fillColor: '#ffeeaa', bold: true },
    })
    engine.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'A1',
      comments: [
        {
          id: 'comment-1',
          body: ' Review ',
          authorUserId: 'user-1',
          authorDisplayName: 'Analyst',
          createdAtUnixMs: 1,
        },
      ],
      resolved: true,
      resolvedByUserId: 'user-2',
      resolvedAtUnixMs: 2,
    })
    engine.setNote({ sheetName: 'Sheet1', address: 'B2', text: ' Note text ' })
    engine.setTable({
      name: 'SalesTable',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
      columnNames: ['Region', 'Sales', 'Units'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setSpillRange('Sheet1', 'E1', 2, 2)
    engine.setPivotTable('Sheet1', 'H1', {
      name: 'Pivot1',
      source: protectedRange,
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Total Sales' }],
    })
    engine.setChart({
      id: 'chart-1',
      sheetName: 'Sheet1',
      address: 'J2',
      source: protectedRange,
      chartType: 'line',
      rows: 8,
      cols: 5,
      seriesOrientation: 'columns',
      firstRowAsHeaders: true,
      firstColumnAsLabels: true,
      title: 'Sales',
      legendPosition: 'right',
    })
    engine.setImage({
      id: 'image-1',
      sheetName: 'Sheet1',
      address: 'L2',
      sourceUrl: 'https://example.com/image.png',
      rows: 4,
      cols: 5,
      altText: 'Image',
    })
    engine.setShape({
      id: 'shape-1',
      sheetName: 'Sheet1',
      address: 'M3',
      shapeType: 'textBox',
      rows: 2,
      cols: 4,
      text: 'Callout',
      fillColor: '#ffffff',
      strokeColor: '#000000',
    })

    expect(engine.getFilters('Sheet1')).toHaveLength(1)
    expect(engine.getSorts('Sheet1')).toHaveLength(1)
    expect(engine.getDataValidation('Sheet1', protectedRange)).toBeDefined()
    expect(engine.getConditionalFormat('format-1')).toBeDefined()
    expect(engine.getCommentThread('Sheet1', 'A1')?.comments[0]?.body).toBe('Review')
    expect(engine.getNote('Sheet1', 'B2')?.text).toBe('Note text')
    expect(engine.getTable('SalesTable')).toBeDefined()
    expect(engine.getSpillRanges()).toHaveLength(1)
    expect(engine.getPivotTable('Sheet1', 'H1')).toBeDefined()
    expect(engine.getChart('chart-1')).toBeDefined()
    expect(engine.getImage('image-1')).toBeDefined()
    expect(engine.getShape('shape-1')).toBeDefined()

    expect(engine.clearFilter('Sheet1', protectedRange)).toBe(true)
    expect(engine.clearSort('Sheet1', protectedRange)).toBe(true)
    expect(engine.clearDataValidation('Sheet1', protectedRange)).toBe(true)
    expect(engine.deleteConditionalFormat('format-1')).toBe(true)
    expect(engine.deleteCommentThread('Sheet1', 'A1')).toBe(true)
    expect(engine.deleteNote('Sheet1', 'B2')).toBe(true)
    expect(engine.deleteTable('SalesTable')).toBe(true)
    expect(engine.deleteSpillRange('Sheet1', 'E1')).toBe(true)
    expect(engine.deletePivotTable('Sheet1', 'H1')).toBe(true)
    expect(engine.deleteChart('chart-1')).toBe(true)
    expect(engine.deleteImage('image-1')).toBe(true)
    expect(engine.deleteShape('shape-1')).toBe(true)

    expect(engine.clearFilter('Sheet1', protectedRange)).toBe(false)
    expect(engine.clearSort('Sheet1', protectedRange)).toBe(false)
    expect(engine.clearDataValidation('Sheet1', protectedRange)).toBe(false)
    expect(engine.deleteConditionalFormat('format-1')).toBe(false)
    expect(engine.deleteCommentThread('Sheet1', 'A1')).toBe(false)
    expect(engine.deleteNote('Sheet1', 'B2')).toBe(false)
    expect(engine.deleteTable('SalesTable')).toBe(false)
    expect(engine.deleteSpillRange('Sheet1', 'E1')).toBe(false)
    expect(engine.deletePivotTable('Sheet1', 'H1')).toBe(false)
    expect(engine.deleteChart('chart-1')).toBe(false)
    expect(engine.deleteImage('image-1')).toBe(false)
    expect(engine.deleteShape('shape-1')).toBe(false)

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('blocks protected metadata mutations across range-backed objects', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-protected-metadata-coverage' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeProtection({ id: 'protect-a1-b2', range: protectedRange })

    expect(() => engine.setFilter('Sheet1', protectedRange)).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setSort('Sheet1', protectedRange, [{ keyAddress: 'A1', direction: 'asc' }])).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() =>
      engine.setDataValidation({
        id: 'validation-1',
        range: protectedRange,
        rule: { kind: 'list', values: ['A', 'B'] },
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.setConditionalFormat({
        id: 'format-1',
        range: protectedRange,
        rule: { kind: 'textContains', text: 'A' },
        style: { italic: true },
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() => engine.setCommentThread({ threadId: 'thread-1', sheetName: 'Sheet1', address: 'A1', comments: [] })).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() => engine.setNote({ sheetName: 'Sheet1', address: 'A1', text: 'Protected' })).toThrow(
      /Workbook protection blocks this change/,
    )
    expect(() =>
      engine.setTable({
        name: 'ProtectedTable',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
        columnNames: ['A', 'B'],
        headerRow: true,
        totalsRow: false,
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.setPivotTable('Sheet1', 'D1', {
        name: 'Pivot1',
        source: protectedRange,
        groupBy: ['A'],
        values: [{ sourceColumn: 'B', summarizeBy: 'sum', outputLabel: 'B' }],
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.setChart({
        id: 'chart-1',
        sheetName: 'Sheet1',
        address: 'D1',
        source: protectedRange,
        chartType: 'bar',
        rows: 4,
        cols: 4,
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.setImage({
        id: 'image-1',
        sheetName: 'Sheet1',
        address: 'A1',
        sourceUrl: 'https://example.com/image.png',
        rows: 4,
        cols: 4,
      }),
    ).toThrow(/Workbook protection blocks this change/)
    expect(() =>
      engine.setShape({
        id: 'shape-1',
        sheetName: 'Sheet1',
        address: 'A1',
        shapeType: 'rectangle',
        rows: 4,
        cols: 4,
      }),
    ).toThrow(/Workbook protection blocks this change/)
  })
})
