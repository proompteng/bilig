import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import type { WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { ValueTag, formatCellDisplayValue } from '@bilig/protocol'
import { applyWorkbookAgentCommandBundleWithUndoCapture } from '../workbook-agent-apply.js'

function createBundle(overrides: Partial<WorkbookAgentCommandBundle> = {}): WorkbookAgentCommandBundle {
  return {
    id: 'bundle-1',
    documentId: 'doc-1',
    threadId: 'thr-1',
    turnId: 'turn-1',
    goalText: 'Populate prepaid expense template',
    summary: 'Write cells in prepaid expenses!A1:I10',
    scope: 'sheet',
    riskClass: 'medium',
    baseRevision: 0,
    createdAtUnixMs: 1,
    context: null,
    commands: [],
    affectedRanges: [],
    estimatedAffectedCells: null,
    ...overrides,
  }
}

describe('workbook agent apply', () => {
  it('captures one undo bundle for multi-cell writeRange commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('prepaid expenses')

    const bundle = createBundle({
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'prepaid expenses',
          startAddress: 'A1',
          values: [
            ['Expense', 'Vendor'],
            ['Insurance', 'Acme'],
          ],
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getCell('prepaid expenses', 'A1').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Expense',
    })
    expect(engine.getCell('prepaid expenses', 'B2').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Acme',
    })
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getCell('prepaid expenses', 'A1').value).toEqual({
      tag: ValueTag.Empty,
    })
    expect(engine.getCell('prepaid expenses', 'B2').value).toEqual({
      tag: ValueTag.Empty,
    })
  })

  it('captures one undo bundle when formatRange stages style and number format together', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 123)

    const bundle = createBundle({
      commands: [
        {
          kind: 'formatRange',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          patch: {
            font: {
              bold: true,
            },
          },
          numberFormat: '0.00',
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getCell('Sheet1', 'A1').format).toBe('0.00')
    expect(engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)?.font?.bold).toBe(true)
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
  })

  it('infers date formatting for numeric serials written under a date-like header', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')

    const bundle = createBundle({
      commands: [
        {
          kind: 'writeRange',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          values: [['Month'], [46023], [46054]],
        },
      ],
    })

    applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    const firstMonth = engine.getCell('Sheet1', 'A2')
    expect(firstMonth.format).toBeDefined()
    expect(formatCellDisplayValue(firstMonth.value, firstMonth.format)).not.toBe('46023')
  })

  it('captures one undo bundle for setRangeFormulas commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)

    const bundle = createBundle({
      commands: [
        {
          kind: 'setRangeFormulas',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B1',
            endAddress: 'B2',
          },
          formulas: [['=A1*2'], ['=A2*2']],
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('A1*2')
    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*2')
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getCell('Sheet1', 'B1').value).toEqual({
      tag: ValueTag.Empty,
    })
    expect(engine.getCell('Sheet1', 'B2').value).toEqual({
      tag: ValueTag.Empty,
    })
  })

  it('infers date formatting for formula results when a date-like header sits above the written range', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'End Date')

    const bundle = createBundle({
      commands: [
        {
          kind: 'setRangeFormulas',
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A2',
            endAddress: 'A3',
          },
          formulas: [['=DATE(2026,1,1)'], ['=DATE(2026,2,1)']],
        },
      ],
    })

    applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    const firstEndDate = engine.getCell('Sheet1', 'A2')
    expect(firstEndDate.format).toBeDefined()
    if (firstEndDate.value.tag !== ValueTag.Number) {
      throw new Error('Expected numeric date serial result')
    }
    expect(formatCellDisplayValue(firstEndDate.value, firstEndDate.format)).not.toBe(String(firstEndDate.value.value))
  })

  it('captures undo for structural and cell commands in the same bundle', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()

    const bundle = createBundle({
      commands: [
        {
          kind: 'createSheet',
          name: 'prepaid expenses',
        },
        {
          kind: 'writeRange',
          sheetName: 'prepaid expenses',
          startAddress: 'A1',
          values: [['Expense']],
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toContain('prepaid expenses')
    expect(engine.getCell('prepaid expenses', 'A1').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Expense',
    })
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).not.toContain('prepaid expenses')
  })

  it('captures undo for row and column structural commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Header')
    engine.setCellValue('Sheet1', 'A2', 'Value')
    engine.setCellValue('Sheet1', 'B1', 'Extra')

    const bundle = createBundle({
      commands: [
        {
          kind: 'insertRows',
          sheetName: 'Sheet1',
          start: 1,
          count: 1,
        },
        {
          kind: 'deleteColumns',
          sheetName: 'Sheet1',
          start: 1,
          count: 1,
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getCell('Sheet1', 'A3').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Value',
    })
    expect(engine.getCell('Sheet1', 'B1').value).toEqual({
      tag: ValueTag.Empty,
    })
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getCell('Sheet1', 'A2').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Value',
    })
    expect(engine.getCell('Sheet1', 'B1').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Extra',
    })
  })

  it('captures undo for deleting a sheet', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Imports')
    engine.setCellValue('Imports', 'A1', 'Raw')

    const bundle = createBundle({
      commands: [
        {
          kind: 'deleteSheet',
          name: 'Imports',
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).not.toContain('Imports')
    expect(undoBundle).toEqual(
      expect.objectContaining({
        kind: 'engineOps',
      }),
    )
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toContain('Imports')
    expect(engine.getCell('Imports', 'A1').value).toMatchObject({
      tag: ValueTag.String,
      value: 'Raw',
    })
  })

  it('captures undo for freeze, filter, and sort metadata commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Header')
    engine.setCellValue('Sheet1', 'A2', 'Value')

    const setBundle = createBundle({
      commands: [
        {
          kind: 'setFreezePane',
          sheetName: 'Sheet1',
          rows: 1,
          cols: 1,
        },
        {
          kind: 'setFilter',
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        },
        {
          kind: 'setSort',
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
          keys: [{ keyAddress: 'B1', direction: 'asc' }],
        },
      ],
    })

    const setUndoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, setBundle)

    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })
    expect(engine.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      },
    ])
    expect(engine.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    if (!setUndoBundle || setUndoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    const clearBundle = createBundle({
      commands: [
        {
          kind: 'setFreezePane',
          sheetName: 'Sheet1',
          rows: 0,
          cols: 0,
        },
        {
          kind: 'clearFilter',
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        },
        {
          kind: 'clearSort',
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        },
      ],
    })

    const clearUndoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, clearBundle)

    expect(engine.getFreezePane('Sheet1')).toBeUndefined()
    expect(engine.getFilters('Sheet1')).toEqual([])
    expect(engine.getSorts('Sheet1')).toEqual([])
    if (!clearUndoBundle || clearUndoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(clearUndoBundle.ops, { trusted: true })

    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })
    expect(engine.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      },
    ])
    expect(engine.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])

    engine.applyOps(setUndoBundle.ops, { trusted: true })

    expect(engine.getFreezePane('Sheet1')).toBeUndefined()
    expect(engine.getFilters('Sheet1')).toEqual([])
    expect(engine.getSorts('Sheet1')).toEqual([])
  })

  it('captures undo for named ranges, tables, and pivots', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Revenue')
    engine.setCellValue('Sheet1', 'B1', 'Margin')
    engine.setCellValue('Sheet1', 'A2', 10)
    engine.setCellValue('Sheet1', 'B2', 2)

    const bundle = createBundle({
      commands: [
        {
          kind: 'upsertDefinedName',
          name: 'Inputs',
          value: {
            kind: 'range-ref',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B2',
          },
        },
        {
          kind: 'upsertTable',
          table: {
            name: 'RevenueTable',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B2',
            columnNames: ['Revenue', 'Margin'],
            headerRow: true,
            totalsRow: false,
          },
        },
        {
          kind: 'upsertPivotTable',
          pivot: {
            name: 'RevenuePivot',
            sheetName: 'Sheet1',
            address: 'E2',
            source: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'B2',
            },
            groupBy: ['Revenue'],
            values: [{ sourceColumn: 'Margin', summarizeBy: 'sum' }],
            rows: 1,
            cols: 2,
          },
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getDefinedName('Inputs')).toBeTruthy()
    expect(engine.getTable('RevenueTable')).toBeTruthy()
    expect(engine.getPivotTable('Sheet1', 'E2')).toBeTruthy()
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getDefinedName('Inputs')).toBeUndefined()
    expect(engine.getTable('RevenueTable')).toBeUndefined()
    expect(engine.getPivotTable('Sheet1', 'E2')).toBeUndefined()
  })

  it('captures undo for data validation commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')

    const bundle = createBundle({
      commands: [
        {
          kind: 'setDataValidation',
          validation: {
            range: {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B4',
            },
            rule: {
              kind: 'list',
              values: ['Draft', 'Final'],
            },
            allowBlank: false,
            showDropdown: true,
          },
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getDataValidations('Sheet1')).toEqual([
      {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
      },
    ])
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getDataValidations('Sheet1')).toEqual([])
  })

  it('captures undo for comment thread and note commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')

    const bundle = createBundle({
      commands: [
        {
          kind: 'upsertCommentThread',
          thread: {
            threadId: 'thread-1',
            sheetName: 'Sheet1',
            address: 'B2',
            comments: [{ id: 'comment-1', body: 'Check this total.' }],
          },
        },
        {
          kind: 'upsertNote',
          note: {
            sheetName: 'Sheet1',
            address: 'C3',
            text: 'Manual override',
          },
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getCommentThreads('Sheet1')).toEqual([
      {
        threadId: 'thread-1',
        sheetName: 'Sheet1',
        address: 'B2',
        comments: [{ id: 'comment-1', body: 'Check this total.' }],
      },
    ])
    expect(engine.getNotes('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        address: 'C3',
        text: 'Manual override',
      },
    ])
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getCommentThreads('Sheet1')).toEqual([])
    expect(engine.getNotes('Sheet1')).toEqual([])
  })

  it('captures undo for conditional format commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')

    const bundle = createBundle({
      commands: [
        {
          kind: 'upsertConditionalFormat',
          format: {
            id: 'cf-1',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B4',
            },
            rule: {
              kind: 'cellIs',
              operator: 'greaterThan',
              values: [10],
            },
            style: {
              fill: { backgroundColor: '#ff0000' },
            },
          },
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getConditionalFormats('Sheet1')).toEqual([
      {
        id: 'cf-1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'cellIs',
          operator: 'greaterThan',
          values: [10],
        },
        style: {
          fill: { backgroundColor: '#ff0000' },
        },
      },
    ])
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getConditionalFormats('Sheet1')).toEqual([])
  })

  it('captures undo for protection commands', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')

    const bundle = createBundle({
      commands: [
        {
          kind: 'setSheetProtection',
          protection: {
            sheetName: 'Sheet1',
            hideFormulas: true,
          },
        },
        {
          kind: 'upsertRangeProtection',
          protection: {
            id: 'protect-a1',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'B2',
            },
            hideFormulas: true,
          },
        },
      ],
    })

    const undoBundle = applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)

    expect(engine.getSheetProtection('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(engine.getRangeProtections('Sheet1')).toEqual([
      {
        id: 'protect-a1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B2',
        },
        hideFormulas: true,
      },
    ])
    if (!undoBundle || undoBundle.kind !== 'engineOps') {
      throw new Error('Expected engineOps undo bundle')
    }

    engine.applyOps(undoBundle.ops, { trusted: true })

    expect(engine.getSheetProtection('Sheet1')).toBeUndefined()
    expect(engine.getRangeProtections('Sheet1')).toEqual([])
  })
})
