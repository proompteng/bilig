import { describe, expect, it } from 'vitest'
import { ValueTag, createCellNumberFormatRecord } from '@bilig/protocol'
import type { StructuralAxisTransform } from '@bilig/formula'
import { writeLiteralToCellStore } from '../engine-value-utils.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'

function hasStructuralAxisTransform(value: unknown): value is {
  applyStructuralAxisTransform: (
    sheetName: string,
    transform: StructuralAxisTransform,
  ) => {
    removedCellIndices: number[]
    remappedCells: Array<{ toRow: number | undefined }>
  }
} {
  return (
    typeof value === 'object' &&
    value !== null &&
    'applyStructuralAxisTransform' in value &&
    typeof value.applyStructuralAxisTransform === 'function'
  )
}

function projectAxisEntryIds(entries: Array<{ id: string; index: number }>): Array<{ id: string; index: number }> {
  return entries.map(({ id, index }) => ({ id, index }))
}

describe('WorkbookStore', () => {
  it('keeps logical cell lookups stable when metadata materializes structural axis entries', () => {
    const workbook = new WorkbookStore('logical-axis-stability')
    workbook.createSheet('Sheet1')

    const cellIndex = workbook.ensureCellRecord('Sheet1', 'B2').cellIndex

    expect(workbook.snapshotRowAxisEntries('Missing', 0, 1)).toEqual([])
    expect(workbook.materializeRowAxisEntries('Sheet1', 0, 0)).toEqual([])

    workbook.setRowMetadata('Sheet1', 1, 1, 30, false)
    workbook.setColumnMetadata('Sheet1', 1, 1, 120, true)

    expect(workbook.getCellIndex('Sheet1', 'B2')).toBe(cellIndex)
    expect(workbook.listRowAxisEntries('Sheet1')).toEqual([{ id: 'row-1', index: 1, size: 30, hidden: false }])
    expect(workbook.listColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 1, size: 120, hidden: true }])
  })

  it('does not mutate existing style ranges when bulk style restoration includes an unknown style', () => {
    const workbook = new WorkbookStore('style-ranges')
    workbook.createSheet('Sheet1')
    workbook.upsertCellStyle({ id: 'style-a', font: { bold: true } })
    workbook.setStyleRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, 'style-a')

    expect(() =>
      workbook.setStyleRanges('Sheet1', [
        {
          range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' },
          styleId: 'style-missing',
        },
      ]),
    ).toThrow('Unknown cell style: style-missing')

    expect(workbook.listStyleRanges('Sheet1')).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        styleId: 'style-a',
      },
    ])
  })

  it('does not mutate existing format ranges when bulk format restoration includes an unknown format', () => {
    const workbook = new WorkbookStore('format-ranges')
    workbook.createSheet('Sheet1')
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-money', '$0.00'))
    workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, 'format-money')

    expect(() =>
      workbook.setFormatRanges('Sheet1', [
        {
          range: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' },
          formatId: 'format-missing',
        },
      ]),
    ).toThrow('Unknown cell number format: format-missing')

    expect(workbook.listFormatRanges('Sheet1')).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
        formatId: 'format-money',
      },
    ])
  })

  it('normalizes filter and sort ranges so equivalent reversed bounds reuse the same record', () => {
    const workbook = new WorkbookStore('normalized-ranges')
    workbook.createSheet('Sheet1')
    const reversedRange = {
      sheetName: 'Sheet1',
      startAddress: 'C3',
      endAddress: 'A1',
    } as const
    const normalizedRange = {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
    } as const

    workbook.setFilter('Sheet1', reversedRange)
    workbook.setFilter('Sheet1', normalizedRange)
    workbook.setSort('Sheet1', reversedRange, [{ keyAddress: 'B1', direction: 'asc' }])
    workbook.setSort('Sheet1', normalizedRange, [{ keyAddress: 'B1', direction: 'desc' }])

    expect(workbook.listFilters('Sheet1')).toEqual([{ sheetName: 'Sheet1', range: normalizedRange }])
    expect(workbook.getFilter('Sheet1', reversedRange)).toEqual({
      sheetName: 'Sheet1',
      range: normalizedRange,
    })
    expect(workbook.deleteFilter('Sheet1', reversedRange)).toBe(true)
    expect(workbook.listFilters('Sheet1')).toEqual([])

    expect(workbook.listSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: normalizedRange,
        keys: [{ keyAddress: 'B1', direction: 'desc' }],
      },
    ])
    expect(workbook.getSort('Sheet1', reversedRange)).toEqual({
      sheetName: 'Sheet1',
      range: normalizedRange,
      keys: [{ keyAddress: 'B1', direction: 'desc' }],
    })
    expect(workbook.deleteSort('Sheet1', reversedRange)).toBe(true)
    expect(workbook.listSorts('Sheet1')).toEqual([])
  })

  it('normalizes data validation ranges so equivalent reversed bounds reuse the same record', () => {
    const workbook = new WorkbookStore('normalized-data-validations')
    workbook.createSheet('Sheet1')
    const reversedRange = {
      sheetName: 'Sheet1',
      startAddress: 'C3',
      endAddress: 'A1',
    } as const
    const normalizedRange = {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C3',
    } as const

    workbook.setDataValidation({
      range: reversedRange,
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
    })
    workbook.setDataValidation({
      range: normalizedRange,
      rule: {
        kind: 'list',
        values: ['Live', 'Archived'],
      },
      allowBlank: true,
    })

    expect(workbook.listDataValidations('Sheet1')).toEqual([
      {
        range: normalizedRange,
        rule: {
          kind: 'list',
          values: ['Live', 'Archived'],
        },
        allowBlank: true,
      },
    ])
    expect(workbook.getDataValidation('Sheet1', reversedRange)).toEqual({
      range: normalizedRange,
      rule: {
        kind: 'list',
        values: ['Live', 'Archived'],
      },
      allowBlank: true,
    })
    expect(workbook.deleteDataValidation('Sheet1', reversedRange)).toBe(true)
    expect(workbook.listDataValidations('Sheet1')).toEqual([])
  })

  it('normalizes comment-thread and note addresses so equivalent case variants reuse the same record', () => {
    const workbook = new WorkbookStore('normalized-annotations')
    workbook.createSheet('Sheet1')

    workbook.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'b2',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })
    workbook.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'B2',
      comments: [{ id: 'comment-1', body: 'Updated.' }],
    })
    workbook.setNote({
      sheetName: 'Sheet1',
      address: 'c3',
      text: 'Manual override',
    })
    workbook.setNote({
      sheetName: 'Sheet1',
      address: 'C3',
      text: 'Updated note',
    })

    expect(workbook.listCommentThreads('Sheet1')).toEqual([
      {
        threadId: 'thread-1',
        sheetName: 'Sheet1',
        address: 'B2',
        comments: [{ id: 'comment-1', body: 'Updated.' }],
      },
    ])
    expect(workbook.listNotes('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        address: 'C3',
        text: 'Updated note',
      },
    ])
    expect(workbook.deleteCommentThread('Sheet1', 'b2')).toBe(true)
    expect(workbook.deleteNote('Sheet1', 'c3')).toBe(true)
    expect(workbook.listCommentThreads('Sheet1')).toEqual([])
    expect(workbook.listNotes('Sheet1')).toEqual([])
  })

  it('reuses conditional format ids while preserving normalized ranges', () => {
    const workbook = new WorkbookStore('conditional-formats')
    workbook.createSheet('Sheet1')

    workbook.setConditionalFormat({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'A1',
      },
      rule: {
        kind: 'cellIs',
        operator: 'greaterThan',
        values: [10],
      },
      style: {
        fill: { backgroundColor: '#ff0000' },
      },
    })
    workbook.setConditionalFormat({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C3',
      },
      rule: {
        kind: 'textContains',
        text: 'urgent',
      },
      style: {
        font: { bold: true },
      },
    })

    expect(workbook.listConditionalFormats('Sheet1')).toEqual([
      {
        id: 'cf-1',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'C3',
        },
        rule: {
          kind: 'textContains',
          text: 'urgent',
        },
        style: {
          font: { bold: true },
        },
      },
    ])
    expect(workbook.getConditionalFormat('cf-1')).toEqual({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C3',
      },
      rule: {
        kind: 'textContains',
        text: 'urgent',
      },
      style: {
        font: { bold: true },
      },
    })
    expect(workbook.deleteConditionalFormat('cf-1')).toBe(true)
    expect(workbook.listConditionalFormats('Sheet1')).toEqual([])
  })

  it('stores sheet and range protections with normalized keys and ranges', () => {
    const workbook = new WorkbookStore('protections')
    workbook.createSheet('Sheet1')

    workbook.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    workbook.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'C3',
        endAddress: 'A1',
      },
      hideFormulas: true,
    })

    expect(workbook.getSheetProtection('Sheet1')).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(workbook.getRangeProtection('protect-a1')).toEqual({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C3',
      },
      hideFormulas: true,
    })
    expect(workbook.listRangeProtections('Sheet1')).toHaveLength(1)
    expect(workbook.clearSheetProtection('Sheet1')).toBe(true)
    expect(workbook.deleteRangeProtection('protect-a1')).toBe(true)
    expect(workbook.getSheetProtection('Sheet1')).toBeUndefined()
    expect(workbook.listRangeProtections('Sheet1')).toEqual([])
  })

  it('normalizes spill and pivot addresses so case-only variants reuse the same record', () => {
    const workbook = new WorkbookStore('normalized-addresses')
    workbook.createSheet('Sheet1')

    workbook.setSpill('Sheet1', 'b2', 2, 3)
    workbook.setSpill('Sheet1', 'B2', 4, 1)

    expect(workbook.listSpills()).toEqual([{ sheetName: 'Sheet1', address: 'B2', rows: 4, cols: 1 }])
    expect(workbook.getSpill('Sheet1', 'b2')).toEqual({
      sheetName: 'Sheet1',
      address: 'B2',
      rows: 4,
      cols: 1,
    })
    expect(workbook.deleteSpill('Sheet1', 'b2')).toBe(true)
    expect(workbook.listSpills()).toEqual([])

    workbook.setPivot({
      name: ' RevenuePivot ',
      sheetName: 'Sheet1',
      address: 'c3',
      source: { sheetName: 'Data', startAddress: 'a1', endAddress: 'b4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })
    workbook.setPivot({
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'C3',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'count' }],
      rows: 4,
      cols: 2,
    })

    expect(workbook.listPivots()).toEqual([
      {
        name: 'RevenuePivot',
        sheetName: 'Sheet1',
        address: 'C3',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ field: 'Sales', summarizeBy: 'count' }],
        rows: 4,
        cols: 2,
      },
    ])
    expect(workbook.getPivot('Sheet1', 'c3')).toEqual({
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'C3',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'count' }],
      rows: 4,
      cols: 2,
    })
    expect(workbook.deletePivot('Sheet1', 'c3')).toBe(true)
    expect(workbook.listPivots()).toEqual([])
  })

  it('renames sheet-scoped metadata through the store without leaving stale keys behind', () => {
    const workbook = new WorkbookStore('rename-metadata')
    workbook.createSheet('Source')
    workbook.setFreezePane('Source', 1, 2)
    workbook.setFilter('Source', {
      sheetName: 'Source',
      startAddress: 'C3',
      endAddress: 'A1',
    })
    workbook.setSpill('Source', 'b2', 2, 3)
    workbook.setPivot({
      name: 'RevenuePivot',
      sheetName: 'Source',
      address: 'c3',
      source: { sheetName: 'Source', startAddress: 'b4', endAddress: 'a1' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })

    expect(workbook.renameSheet('Source', 'Renamed')?.name).toBe('Renamed')

    expect(workbook.getFreezePane('Source')).toBeUndefined()
    expect(workbook.getFreezePane('Renamed')).toEqual({
      sheetName: 'Renamed',
      rows: 1,
      cols: 2,
    })
    expect(
      workbook.getFilter('Renamed', {
        sheetName: 'Renamed',
        startAddress: 'A1',
        endAddress: 'C3',
      }),
    ).toEqual({
      sheetName: 'Renamed',
      range: { sheetName: 'Renamed', startAddress: 'A1', endAddress: 'C3' },
    })
    expect(workbook.getSpill('Renamed', 'B2')).toEqual({
      sheetName: 'Renamed',
      address: 'B2',
      rows: 2,
      cols: 3,
    })
    expect(workbook.getPivot('Renamed', 'C3')).toEqual({
      name: 'RevenuePivot',
      sheetName: 'Renamed',
      address: 'C3',
      source: { sheetName: 'Renamed', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })
  })

  it('does not clear a remapped live cell when pruning an empty removed structural cell', () => {
    const workbook = new WorkbookStore('prune-remapped-cell')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const removedCellIndex = workbook.ensureCell('Sheet1', 'B3')
    const movedCellIndex = workbook.ensureCell('Sheet1', 'B4')
    writeLiteralToCellStore(workbook.cellStore, removedCellIndex, 7, strings)
    writeLiteralToCellStore(workbook.cellStore, movedCellIndex, 5, strings)

    workbook.deleteRows('Sheet1', 2, 1)
    workbook.remapSheetCells('Sheet1', 'row', (index) => (index < 2 ? index : index >= 3 ? index - 1 : undefined))

    workbook.cellStore.setValue(removedCellIndex, { tag: ValueTag.Empty })

    expect(workbook.pruneCellIfEmpty(removedCellIndex)).toBe(false)
    expect(workbook.getCellIndex('Sheet1', 'B3')).toBe(movedCellIndex)
    expect(workbook.cellStore.getValue(movedCellIndex, () => '')).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
  })

  it('restores sparse column axis entries without synthesizing blank identities', () => {
    const workbook = new WorkbookStore('sparse-axis-restore')
    workbook.createSheet('Sheet1')
    workbook.setColumnMetadata('Sheet1', 1, 1, 120, false)

    const captured = workbook.snapshotColumnAxisEntries('Sheet1', 0, 2)
    expect(captured).toEqual([{ id: 'column-1', index: 1, size: 120, hidden: false }])

    workbook.deleteColumns('Sheet1', 0, 2)
    workbook.insertColumns('Sheet1', 0, 2, captured)

    expect(workbook.listColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 1, size: 120, hidden: false }])
  })

  it('mirrors inserted, moved, and deleted column axis ids in lockstep with workbook snapshots', () => {
    const workbook = new WorkbookStore('axis-map-mirror')
    workbook.createSheet('Sheet1')
    const sheet = workbook.getSheet('Sheet1')

    expect(sheet).toBeDefined()
    if (!sheet) {
      throw new Error('Expected Sheet1 to exist')
    }

    workbook.insertColumns('Sheet1', 0, 3, [
      { id: 'column-a', index: 0, size: 120, hidden: false },
      { id: 'column-b', index: 1, size: 90, hidden: true },
      { id: 'column-c', index: 2, size: null, hidden: null },
    ])

    expect(workbook.snapshotColumnAxisEntries('Sheet1', 0, 3)).toEqual([
      { id: 'column-a', index: 0, size: 120, hidden: false },
      { id: 'column-b', index: 1, size: 90, hidden: true },
      { id: 'column-c', index: 2 },
    ])
    expect(sheet.axisMap.list('column')).toEqual([
      { id: 'column-a', index: 0 },
      { id: 'column-b', index: 1 },
      { id: 'column-c', index: 2 },
    ])
    expect(sheet.axisMap.snapshot('column', 0, 3)).toEqual(projectAxisEntryIds(workbook.snapshotColumnAxisEntries('Sheet1', 0, 3)))

    workbook.moveColumns('Sheet1', 0, 1, 3)

    expect(workbook.snapshotColumnAxisEntries('Sheet1', 0, 3)).toEqual([
      { id: 'column-b', index: 0, size: 90, hidden: true },
      { id: 'column-c', index: 1 },
      { id: 'column-a', index: 2, size: 120, hidden: false },
    ])
    expect(sheet.axisMap.list('column')).toEqual([
      { id: 'column-b', index: 0 },
      { id: 'column-c', index: 1 },
      { id: 'column-a', index: 2 },
    ])
    expect(sheet.axisMap.snapshot('column', 0, 3)).toEqual(projectAxisEntryIds(workbook.snapshotColumnAxisEntries('Sheet1', 0, 3)))

    const deleted = workbook.deleteColumns('Sheet1', 1, 1)

    expect(deleted).toEqual([{ id: 'column-c', index: 1 }])
    expect(workbook.snapshotColumnAxisEntries('Sheet1', 0, 2)).toEqual([
      { id: 'column-b', index: 0, size: 90, hidden: true },
      { id: 'column-a', index: 1, size: 120, hidden: false },
    ])
    expect(sheet.axisMap.list('column')).toEqual([
      { id: 'column-b', index: 0 },
      { id: 'column-a', index: 1 },
    ])
    expect(sheet.axisMap.snapshot('column', 0, 2)).toEqual(projectAxisEntryIds(workbook.snapshotColumnAxisEntries('Sheet1', 0, 2)))
  })

  it('remaps only affected cells during structural column shifts', () => {
    const workbook = new WorkbookStore('remap-affected-columns')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const leftCellIndex = workbook.ensureCell('Sheet1', 'A1')
    const movedCellIndex = workbook.ensureCell('Sheet1', 'B1')
    const farCellIndex = workbook.ensureCell('Sheet1', 'D1')
    writeLiteralToCellStore(workbook.cellStore, leftCellIndex, 1, strings)
    writeLiteralToCellStore(workbook.cellStore, movedCellIndex, 2, strings)
    writeLiteralToCellStore(workbook.cellStore, farCellIndex, 4, strings)

    const remapped = workbook.remapSheetCells('Sheet1', 'column', (index) => (index < 1 ? index : index + 1))

    expect(remapped.changedCellIndices).toEqual([movedCellIndex, farCellIndex])
    expect(workbook.getCellIndex('Sheet1', 'A1')).toBe(leftCellIndex)
    expect(workbook.getCellIndex('Sheet1', 'C1')).toBe(movedCellIndex)
    expect(workbook.getCellIndex('Sheet1', 'E1')).toBe(farCellIndex)
    expect(workbook.getCellIndex('Sheet1', 'B1')).toBeUndefined()
    expect(workbook.cellStore.cols[leftCellIndex]).toBe(0)
    expect(workbook.cellStore.cols[movedCellIndex]).toBe(2)
    expect(workbook.cellStore.cols[farCellIndex]).toBe(4)
  })

  it('follows axis-map row inserts through the logical store before physical remap runs', () => {
    const workbook = new WorkbookStore('logical-axis-row-insert')
    workbook.createSheet('Sheet1')

    const cellIndex = workbook.ensureCell('Sheet1', 'A1')

    expect(workbook.getCellIndex('Sheet1', 'A1')).toBe(cellIndex)

    workbook.insertRows('Sheet1', 0, 1)

    expect(workbook.getCellIndex('Sheet1', 'A1')).toBeUndefined()
    expect(workbook.getCellIndex('Sheet1', 'A2')).toBe(cellIndex)
  })

  it('builds one structural transaction from workbook remaps and tracks removed cells', () => {
    const counters = createEngineCounters()
    const workbook = new WorkbookStore('structural-transaction', counters)
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const removedCellIndex = workbook.ensureCell('Sheet1', 'C11')
    const movedCellIndex = workbook.ensureCell('Sheet1', 'C12')
    writeLiteralToCellStore(workbook.cellStore, removedCellIndex, 7, strings)
    writeLiteralToCellStore(workbook.cellStore, movedCellIndex, 8, strings)

    expect(hasStructuralAxisTransform(workbook)).toBe(true)

    const transaction = hasStructuralAxisTransform(workbook)
      ? workbook.applyStructuralAxisTransform('Sheet1', {
          axis: 'row',
          kind: 'delete',
          start: 10,
          count: 1,
        })
      : undefined

    expect(transaction).toBeDefined()
    expect(transaction?.removedCellIndices).toEqual([removedCellIndex])
    expect(transaction?.remappedCells.some((entry) => entry.toRow === 10)).toBe(true)
    expect(workbook.getCellIndex('Sheet1', 'C11')).toBe(movedCellIndex)
    expect(counters.cellsRemapped).toBeGreaterThan(0)
  })
})
