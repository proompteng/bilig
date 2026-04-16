/* eslint-disable typescript-eslint/no-unsafe-type-assertion -- bucket-failure tests intentionally replace strongly typed maps with throwing doubles */
import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { createWorkbookMetadataService, runWorkbookMetadataEffect, type WorkbookMetadataService } from '../workbook-metadata-service.js'
import { createWorkbookMetadataRecord } from '../workbook-metadata-types.js'

function createService(): WorkbookMetadataService {
  return createWorkbookMetadataService(createWorkbookMetadataRecord())
}

function createThrowingBucket(message: string): Map<string, unknown> {
  const fail = () => {
    throw new Error(message)
  }
  return {
    set: fail,
    get: fail,
    delete: fail,
    values: fail,
    clear: fail,
  } as unknown as Map<string, unknown>
}

describe('WorkbookMetadataService', () => {
  it('clones caller-owned defined name objects on write and read', () => {
    const service = createService()
    const source = {
      kind: 'range-ref' as const,
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
    }

    const stored = Effect.runSync(service.setDefinedName(' SalesRange ', source))
    source.sheetName = 'Mutated'
    if (stored.value && typeof stored.value === 'object' && stored.value.kind === 'range-ref') {
      stored.value.startAddress = 'Z9'
    }

    expect(Effect.runSync(service.getDefinedName('salesrange'))).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName: 'Data',
        startAddress: 'A1',
        endAddress: 'B4',
      },
    })
    expect(Effect.runSync(service.listDefinedNames())).toEqual([
      {
        name: 'SalesRange',
        value: {
          kind: 'range-ref',
          sheetName: 'Data',
          startAddress: 'A1',
          endAddress: 'B4',
        },
      },
    ])
  })

  it('clones and normalizes data validation records on write and read', () => {
    const service = createService()
    const input = {
      range: {
        sheetName: 'Sheet1',
        startAddress: 'c4',
        endAddress: 'b2',
      },
      rule: {
        kind: 'list' as const,
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: 'stop' as const,
      errorTitle: 'Status required',
      errorMessage: 'Pick Draft or Final.',
    }

    const stored = Effect.runSync(service.setDataValidation(input))
    input.rule.values[0] = 'Mutated'
    stored.rule.values.push('Broken')

    expect(
      Effect.runSync(
        service.getDataValidation('Sheet1', {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C4',
        }),
      ),
    ).toEqual({
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: 'stop',
      errorTitle: 'Status required',
      errorMessage: 'Pick Draft or Final.',
    })
  })

  it('clones and normalizes comment threads and notes on write and read', () => {
    const service = createService()
    const threadInput = {
      threadId: ' thread-1 ',
      sheetName: 'Sheet1',
      address: 'c4',
      comments: [{ id: ' comment-1 ', body: 'Check this total.' }],
    }
    const noteInput = {
      sheetName: 'Sheet1',
      address: 'd5',
      text: ' Manual override ',
    }

    const storedThread = Effect.runSync(service.setCommentThread(threadInput))
    const storedNote = Effect.runSync(service.setNote(noteInput))
    Effect.runSync(service.setNote({ sheetName: 'Sheet1', address: 'B1', text: 'Earlier note' }))
    threadInput.comments[0].body = 'Mutated'
    storedThread.comments[0].body = 'Leaked'
    noteInput.text = 'Changed'
    storedNote.text = 'Broken'

    expect(Effect.runSync(service.getCommentThread('Sheet1', 'C4'))).toEqual({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'C4',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })
    expect(Effect.runSync(service.getNote('Sheet1', 'D5'))).toEqual({
      sheetName: 'Sheet1',
      address: 'D5',
      text: 'Manual override',
    })
    expect(Effect.runSync(service.listNotes('Sheet1'))).toEqual([
      { sheetName: 'Sheet1', address: 'B1', text: 'Earlier note' },
      { sheetName: 'Sheet1', address: 'D5', text: 'Manual override' },
    ])
  })

  it('clones and normalizes conditional formats on write and read', () => {
    const service = createService()
    const input = {
      id: ' cf-1 ',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'c4',
        endAddress: 'b2',
      },
      rule: {
        kind: 'cellIs' as const,
        operator: 'greaterThan' as const,
        values: [10],
      },
      style: {
        fill: { backgroundColor: '#ff0000' },
      },
      stopIfTrue: true,
      priority: 1,
    }

    const stored = Effect.runSync(service.setConditionalFormat(input))
    input.rule.values[0] = 99
    stored.style.fill!.backgroundColor = '#00ff00'

    expect(Effect.runSync(service.getConditionalFormat('cf-1'))).toEqual({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
      rule: {
        kind: 'cellIs',
        operator: 'greaterThan',
        values: [10],
      },
      style: {
        fill: { backgroundColor: '#ff0000' },
      },
      stopIfTrue: true,
      priority: 1,
    })
  })

  it('clones and normalizes sheet and range protection records on write and read', () => {
    const service = createService()
    const sheet = Effect.runSync(
      service.setSheetProtection({
        sheetName: 'Sheet1',
        hideFormulas: true,
      }),
    )
    const range = Effect.runSync(
      service.setRangeProtection({
        id: ' protect-a1 ',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'c4',
          endAddress: 'b2',
        },
        hideFormulas: true,
      }),
    )
    sheet.hideFormulas = false
    range.range.startAddress = 'Z9'

    expect(Effect.runSync(service.getSheetProtection('Sheet1'))).toEqual({
      sheetName: 'Sheet1',
      hideFormulas: true,
    })
    expect(Effect.runSync(service.getRangeProtection('protect-a1'))).toEqual({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'C4',
      },
      hideFormulas: true,
    })
  })

  it('normalizes and clones pivot records so caller mutation does not leak back into metadata', () => {
    const service = createService()
    const input = {
      name: ' RevenuePivot ',
      sheetName: 'Sheet1',
      address: 'c3',
      source: { sheetName: 'Data', startAddress: 'b4', endAddress: 'a1' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' as const }],
      rows: 3,
      cols: 2,
    }

    const stored = Effect.runSync(service.setPivot(input))
    input.groupBy[0] = 'Mutated'
    input.values[0].summarizeBy = 'count'
    stored.groupBy.push('Leaked')
    stored.values[0].field = 'Broken'
    stored.source.startAddress = 'Z9'

    expect(Effect.runSync(service.getPivot('Sheet1', 'C3'))).toEqual({
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'C3',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })
    expect(Effect.runSync(service.listPivots())).toEqual([
      {
        name: 'RevenuePivot',
        sheetName: 'Sheet1',
        address: 'C3',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ field: 'Sales', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      },
    ])
  })

  it('renames, deletes, and resets workbook metadata across buckets', () => {
    const metadata = createWorkbookMetadataRecord()
    const service = createWorkbookMetadataService(metadata)

    Effect.runSync(service.setFreezePane('Source', 1, 2))
    Effect.runSync(
      service.setFilter('Source', {
        sheetName: 'Source',
        startAddress: 'C3',
        endAddress: 'A1',
      }),
    )
    Effect.runSync(
      service.setSort('Source', { sheetName: 'Source', startAddress: 'C3', endAddress: 'A1' }, [{ keyAddress: 'B1', direction: 'asc' }]),
    )
    Effect.runSync(
      service.setDataValidation({
        range: { sheetName: 'Source', startAddress: 'B4', endAddress: 'A2' },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
      }),
    )
    Effect.runSync(
      service.setCommentThread({
        threadId: 'thread-1',
        sheetName: 'Source',
        address: 'C4',
        comments: [{ id: 'comment-1', body: 'Check this total.' }],
      }),
    )
    Effect.runSync(
      service.setNote({
        sheetName: 'Source',
        address: 'D5',
        text: 'Manual override',
      }),
    )
    Effect.runSync(
      service.setConditionalFormat({
        id: 'cf-1',
        range: { sheetName: 'Source', startAddress: 'A2', endAddress: 'A5' },
        rule: {
          kind: 'cellIs',
          operator: 'greaterThan',
          values: [10],
        },
        style: {
          fill: { backgroundColor: '#ff0000' },
        },
      }),
    )
    Effect.runSync(
      service.setSheetProtection({
        sheetName: 'Source',
        hideFormulas: true,
      }),
    )
    Effect.runSync(
      service.setRangeProtection({
        id: 'protect-a1',
        range: { sheetName: 'Source', startAddress: 'A2', endAddress: 'A5' },
        hideFormulas: true,
      }),
    )
    Effect.runSync(
      service.setTable({
        name: ' Revenue ',
        sheetName: 'Source',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      }),
    )
    Effect.runSync(service.setSpill('Source', 'b2', 2, 3))
    Effect.runSync(
      service.setPivot({
        name: ' RevenuePivot ',
        sheetName: 'Source',
        address: 'c3',
        source: { sheetName: 'Source', startAddress: 'b4', endAddress: 'a1' },
        groupBy: ['Region'],
        values: [{ field: 'Sales', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      }),
    )
    Effect.runSync(service.setWorkbookProperty(' Author ', 'greg'))
    Effect.runSync(service.setDefinedName(' LocalName ', { kind: 'formula', formula: '=Source!A1' }))
    Effect.runSync(service.setCalculationSettings({ mode: 'manual' }))
    Effect.runSync(service.setVolatileContext({ recalcEpoch: 7 }))
    metadata.rowMetadata.set('Source:0:2', {
      sheetName: 'Source',
      start: 0,
      count: 2,
      size: 24,
      hidden: null,
    })
    metadata.columnMetadata.set('Source:1:1', {
      sheetName: 'Source',
      start: 1,
      count: 1,
      size: null,
      hidden: true,
    })

    Effect.runSync(service.renameSheet('Source', 'Renamed'))

    expect(Effect.runSync(service.getFreezePane('Renamed'))).toEqual({
      sheetName: 'Renamed',
      rows: 1,
      cols: 2,
    })
    expect(
      Effect.runSync(
        service.getFilter('Renamed', {
          sheetName: 'Renamed',
          startAddress: 'A1',
          endAddress: 'C3',
        }),
      ),
    ).toEqual({
      sheetName: 'Renamed',
      range: { sheetName: 'Renamed', startAddress: 'A1', endAddress: 'C3' },
    })
    expect(
      Effect.runSync(
        service.getSort('Renamed', {
          sheetName: 'Renamed',
          startAddress: 'A1',
          endAddress: 'C3',
        }),
      ),
    ).toEqual({
      sheetName: 'Renamed',
      range: { sheetName: 'Renamed', startAddress: 'A1', endAddress: 'C3' },
      keys: [{ keyAddress: 'B1', direction: 'asc' }],
    })
    expect(
      Effect.runSync(
        service.getDataValidation('Renamed', {
          sheetName: 'Renamed',
          startAddress: 'A2',
          endAddress: 'B4',
        }),
      ),
    ).toEqual({
      range: { sheetName: 'Renamed', startAddress: 'A2', endAddress: 'B4' },
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: false,
      showDropdown: true,
    })
    expect(Effect.runSync(service.getCommentThread('Renamed', 'C4'))).toEqual({
      threadId: 'thread-1',
      sheetName: 'Renamed',
      address: 'C4',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })
    expect(Effect.runSync(service.getNote('Renamed', 'D5'))).toEqual({
      sheetName: 'Renamed',
      address: 'D5',
      text: 'Manual override',
    })
    expect(Effect.runSync(service.getConditionalFormat('cf-1'))).toEqual({
      id: 'cf-1',
      range: { sheetName: 'Renamed', startAddress: 'A2', endAddress: 'A5' },
      rule: {
        kind: 'cellIs',
        operator: 'greaterThan',
        values: [10],
      },
      style: {
        fill: { backgroundColor: '#ff0000' },
      },
    })
    expect(Effect.runSync(service.getSheetProtection('Renamed'))).toEqual({
      sheetName: 'Renamed',
      hideFormulas: true,
    })
    expect(Effect.runSync(service.getRangeProtection('protect-a1'))).toEqual({
      id: 'protect-a1',
      range: { sheetName: 'Renamed', startAddress: 'A2', endAddress: 'A5' },
      hideFormulas: true,
    })
    expect(Effect.runSync(service.getSpill('Renamed', 'B2'))).toEqual({
      sheetName: 'Renamed',
      address: 'B2',
      rows: 2,
      cols: 3,
    })
    expect(Effect.runSync(service.getPivot('Renamed', 'C3'))).toEqual({
      name: 'RevenuePivot',
      sheetName: 'Renamed',
      address: 'C3',
      source: { sheetName: 'Renamed', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })
    expect(Effect.runSync(service.getTable('Revenue'))).toEqual({
      name: 'Revenue',
      sheetName: 'Renamed',
      startAddress: 'A1',
      endAddress: 'C10',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    expect([...metadata.rowMetadata.values()]).toEqual([{ sheetName: 'Renamed', start: 0, count: 2, size: 24, hidden: null }])
    expect([...metadata.columnMetadata.values()]).toEqual([{ sheetName: 'Renamed', start: 1, count: 1, size: null, hidden: true }])

    Effect.runSync(service.deleteSheetRecords('Renamed'))

    expect(Effect.runSync(service.getFreezePane('Renamed'))).toBeUndefined()
    expect(Effect.runSync(service.listFilters('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listSorts('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listDataValidations('Renamed'))).toEqual([])
    expect(Effect.runSync(service.getSheetProtection('Renamed'))).toBeUndefined()
    expect(Effect.runSync(service.listConditionalFormats('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listRangeProtections('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listCommentThreads('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listNotes('Renamed'))).toEqual([])
    expect(Effect.runSync(service.listPivots())).toEqual([])
    expect(Effect.runSync(service.listTables())).toEqual([])
    expect(Effect.runSync(service.listSpills())).toEqual([])
    expect([...metadata.rowMetadata.values()]).toEqual([])
    expect([...metadata.columnMetadata.values()]).toEqual([])
    expect(Effect.runSync(service.getWorkbookProperty('Author'))).toEqual({
      key: 'Author',
      value: 'greg',
    })
    expect(Effect.runSync(service.getDefinedName('LocalName'))).toEqual({
      name: 'LocalName',
      value: { kind: 'formula', formula: '=Source!A1' },
    })

    Effect.runSync(service.reset())

    expect(Effect.runSync(service.listWorkbookProperties())).toEqual([])
    expect(Effect.runSync(service.listDefinedNames())).toEqual([])
    expect(Effect.runSync(service.getCalculationSettings())).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
    })
    expect(Effect.runSync(service.getVolatileContext())).toEqual({
      recalcEpoch: 0,
    })
  })

  it('supports delete and clear helpers across metadata buckets', () => {
    const service = createService()

    expect(Effect.runSync(service.setWorkbookProperty(' Author ', 'greg'))).toEqual({
      key: 'Author',
      value: 'greg',
    })
    expect(Effect.runSync(service.setWorkbookProperty('Author', null))).toBeUndefined()
    expect(Effect.runSync(service.getWorkbookProperty('Author'))).toBeUndefined()

    Effect.runSync(service.setDefinedName(' LocalName ', { kind: 'scalar', value: 1 }))
    expect(Effect.runSync(service.deleteDefinedName('localname'))).toBe(true)
    expect(Effect.runSync(service.deleteDefinedName('localname'))).toBe(false)

    Effect.runSync(
      service.setTable({
        name: ' Revenue ',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B4',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      }),
    )
    expect(Effect.runSync(service.deleteTable('revenue'))).toBe(true)
    expect(Effect.runSync(service.deleteTable('revenue'))).toBe(false)

    Effect.runSync(service.setFreezePane('Sheet1', 1, 1))
    expect(Effect.runSync(service.clearFreezePane('Sheet1'))).toBe(true)
    expect(Effect.runSync(service.clearFreezePane('Sheet1'))).toBe(false)

    Effect.runSync(service.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true }))
    expect(Effect.runSync(service.clearSheetProtection('Sheet1'))).toBe(true)
    expect(Effect.runSync(service.clearSheetProtection('Sheet1'))).toBe(false)

    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    Effect.runSync(service.setFilter('Sheet1', range))
    expect(Effect.runSync(service.deleteFilter('Sheet1', range))).toBe(true)
    expect(Effect.runSync(service.deleteFilter('Sheet1', range))).toBe(false)

    Effect.runSync(service.setSort('Sheet1', range, [{ keyAddress: 'A1', direction: 'asc' }]))
    expect(Effect.runSync(service.deleteSort('Sheet1', range))).toBe(true)
    expect(Effect.runSync(service.deleteSort('Sheet1', range))).toBe(false)

    Effect.runSync(
      service.setDataValidation({
        range,
        rule: { kind: 'list', values: ['Draft', 'Final'] },
      }),
    )
    expect(Effect.runSync(service.deleteDataValidation('Sheet1', range))).toBe(true)
    expect(Effect.runSync(service.deleteDataValidation('Sheet1', range))).toBe(false)

    Effect.runSync(
      service.setConditionalFormat({
        id: 'cf-1',
        range,
        rule: { kind: 'blanks' },
        style: {},
      }),
    )
    expect(Effect.runSync(service.deleteConditionalFormat('cf-1'))).toBe(true)
    expect(Effect.runSync(service.deleteConditionalFormat('cf-1'))).toBe(false)

    Effect.runSync(
      service.setRangeProtection({
        id: 'protect-1',
        range,
      }),
    )
    expect(Effect.runSync(service.deleteRangeProtection('protect-1'))).toBe(true)
    expect(Effect.runSync(service.deleteRangeProtection('protect-1'))).toBe(false)

    Effect.runSync(
      service.setCommentThread({
        threadId: 'thread-1',
        sheetName: 'Sheet1',
        address: 'A1',
        comments: [{ id: 'comment-1', body: 'Check this total.' }],
      }),
    )
    expect(Effect.runSync(service.deleteCommentThread('Sheet1', 'A1'))).toBe(true)
    expect(Effect.runSync(service.deleteCommentThread('Sheet1', 'A1'))).toBe(false)

    Effect.runSync(service.setNote({ sheetName: 'Sheet1', address: 'A1', text: 'Watch this' }))
    expect(Effect.runSync(service.deleteNote('Sheet1', 'A1'))).toBe(true)
    expect(Effect.runSync(service.deleteNote('Sheet1', 'A1'))).toBe(false)

    Effect.runSync(service.setSpill('Sheet1', 'A1', 2, 2))
    expect(Effect.runSync(service.deleteSpill('Sheet1', 'A1'))).toBe(true)
    expect(Effect.runSync(service.deleteSpill('Sheet1', 'A1'))).toBe(false)

    Effect.runSync(
      service.setPivot({
        name: 'RevenuePivot',
        sheetName: 'Sheet1',
        address: 'C3',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
        groupBy: ['Region'],
        values: [{ field: 'Sales', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      }),
    )
    expect(Effect.runSync(service.getPivotByKey('Sheet1!C3'))).toEqual({
      name: 'RevenuePivot',
      sheetName: 'Sheet1',
      address: 'C3',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ field: 'Sales', summarizeBy: 'sum' }],
      rows: 3,
      cols: 2,
    })
    expect(Effect.runSync(service.deletePivot('Sheet1', 'C3'))).toBe(true)
    expect(Effect.runSync(service.deletePivot('Sheet1', 'C3'))).toBe(false)

    expect(Effect.runSync(service.listWorkbookProperties())).toEqual([])
    expect(Effect.runSync(service.listDefinedNames())).toEqual([])
    expect(Effect.runSync(service.listTables())).toEqual([])
    expect(Effect.runSync(service.listFilters('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listSorts('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listDataValidations('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listConditionalFormats('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listRangeProtections('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listCommentThreads('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listNotes('Sheet1'))).toEqual([])
    expect(Effect.runSync(service.listSpills())).toEqual([])
    expect(Effect.runSync(service.listPivots())).toEqual([])

    expect(() => runWorkbookMetadataEffect(service.setWorkbookProperty('   ', 'broken'))).toThrow(
      'Workbook metadata keys must be non-empty',
    )
  })

  it('clones, normalizes, lists, and deletes chart, image, and shape metadata', () => {
    const service = createService()
    const chartInput = {
      id: ' chart-b ',
      sheetName: 'Sheet1',
      address: 'c4',
      source: { sheetName: 'Data', startAddress: 'b4', endAddress: 'a1' },
      chartType: 'column' as const,
      seriesOrientation: 'rows' as const,
      firstRowAsHeaders: true,
      firstColumnAsLabels: false,
      title: 'Revenue',
      legendPosition: 'bottom' as const,
      rows: 5,
      cols: 6,
    }
    const imageInput = {
      id: ' image-a ',
      sheetName: 'Sheet1',
      address: 'd6',
      sourceUrl: 'https://example.com/revenue.png',
      rows: 3,
      cols: 4,
      altText: 'Revenue chart',
    }
    const shapeInput = {
      id: ' shape-c ',
      sheetName: 'Sheet1',
      address: 'e7',
      shapeType: 'textBox' as const,
      rows: 2,
      cols: 3,
      text: 'Quarterly note',
      fillColor: '#ffeeaa',
      strokeColor: '#222222',
    }

    const storedChart = Effect.runSync(service.setChart(chartInput))
    const storedImage = Effect.runSync(service.setImage(imageInput))
    const storedShape = Effect.runSync(service.setShape(shapeInput))

    chartInput.title = 'Mutated'
    chartInput.source.startAddress = 'Z9'
    storedChart.title = 'Broken'
    storedChart.source.startAddress = 'Y8'
    imageInput.altText = 'Mutated'
    storedImage.altText = 'Broken'
    shapeInput.text = 'Mutated'
    storedShape.text = 'Broken'

    expect(Effect.runSync(service.getChart('chart-b'))).toEqual({
      id: 'chart-b',
      sheetName: 'Sheet1',
      address: 'C4',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      chartType: 'column',
      seriesOrientation: 'rows',
      firstRowAsHeaders: true,
      firstColumnAsLabels: false,
      title: 'Revenue',
      legendPosition: 'bottom',
      rows: 5,
      cols: 6,
    })
    expect(Effect.runSync(service.getImage('image-a'))).toEqual({
      id: 'image-a',
      sheetName: 'Sheet1',
      address: 'D6',
      sourceUrl: 'https://example.com/revenue.png',
      rows: 3,
      cols: 4,
      altText: 'Revenue chart',
    })
    expect(Effect.runSync(service.getShape('shape-c'))).toEqual({
      id: 'shape-c',
      sheetName: 'Sheet1',
      address: 'E7',
      shapeType: 'textBox',
      rows: 2,
      cols: 3,
      text: 'Quarterly note',
      fillColor: '#ffeeaa',
      strokeColor: '#222222',
    })

    Effect.runSync(
      service.setChart({
        id: ' chart-a ',
        sheetName: 'Sheet1',
        address: 'a1',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'A2' },
        chartType: 'line',
        rows: 2,
        cols: 2,
      }),
    )
    Effect.runSync(
      service.setImage({
        id: ' image-b ',
        sheetName: 'Sheet1',
        address: 'f1',
        sourceUrl: 'https://example.com/logo.png',
        rows: 1,
        cols: 1,
      }),
    )
    Effect.runSync(
      service.setShape({
        id: ' shape-a ',
        sheetName: 'Sheet1',
        address: 'g1',
        shapeType: 'roundedRectangle',
        rows: 1,
        cols: 2,
      }),
    )

    expect(Effect.runSync(service.listCharts()).map((record) => record.id)).toEqual(['chart-a', 'chart-b'])
    expect(Effect.runSync(service.listImages()).map((record) => record.id)).toEqual(['image-a', 'image-b'])
    expect(Effect.runSync(service.listShapes()).map((record) => record.id)).toEqual(['shape-a', 'shape-c'])

    expect(Effect.runSync(service.deleteChart('chart-b'))).toBe(true)
    expect(Effect.runSync(service.deleteChart('chart-b'))).toBe(false)
    expect(Effect.runSync(service.deleteImage('image-a'))).toBe(true)
    expect(Effect.runSync(service.deleteImage('image-a'))).toBe(false)
    expect(Effect.runSync(service.deleteShape('shape-c'))).toBe(true)
    expect(Effect.runSync(service.deleteShape('shape-c'))).toBe(false)
  })

  it('normalizes and lists spill metadata records', () => {
    const service = createService()

    const first = Effect.runSync(service.setSpill('Sheet2', 'c4', 3, 2))
    const second = Effect.runSync(service.setSpill('Sheet1', 'b2', 1, 4))
    first.address = 'Z9'
    second.rows = 99

    expect(Effect.runSync(service.getSpill('Sheet2', 'C4'))).toEqual({
      sheetName: 'Sheet2',
      address: 'C4',
      rows: 3,
      cols: 2,
    })
    expect(Effect.runSync(service.listSpills())).toEqual([
      {
        sheetName: 'Sheet1',
        address: 'B2',
        rows: 1,
        cols: 4,
      },
      {
        sheetName: 'Sheet2',
        address: 'C4',
        rows: 3,
        cols: 2,
      },
    ])
  })

  it('wraps metadata bucket failures with stable workbook metadata errors', () => {
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    const cases = [
      {
        bucket: 'properties',
        message: 'Failed to set workbook property',
        run: (service: WorkbookMetadataService) => service.setWorkbookProperty('Author', 'greg'),
      },
      {
        bucket: 'properties',
        message: 'Failed to get workbook property',
        run: (service: WorkbookMetadataService) => service.getWorkbookProperty('Author'),
      },
      {
        bucket: 'properties',
        message: 'Failed to list workbook properties',
        run: (service: WorkbookMetadataService) => service.listWorkbookProperties(),
      },
      {
        bucket: 'definedNames',
        message: 'Failed to set defined name',
        run: (service: WorkbookMetadataService) => service.setDefinedName('TaxRate', { kind: 'scalar', value: 1 }),
      },
      {
        bucket: 'definedNames',
        message: 'Failed to get defined name',
        run: (service: WorkbookMetadataService) => service.getDefinedName('TaxRate'),
      },
      {
        bucket: 'definedNames',
        message: 'Failed to delete defined name',
        run: (service: WorkbookMetadataService) => service.deleteDefinedName('TaxRate'),
      },
      {
        bucket: 'tables',
        message: 'Failed to set table metadata',
        run: (service: WorkbookMetadataService) =>
          service.setTable({
            name: 'Revenue',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B2',
            columnNames: ['Region'],
            headerRow: true,
            totalsRow: false,
          }),
      },
      {
        bucket: 'freezePanes',
        message: 'Failed to clear freeze pane metadata',
        run: (service: WorkbookMetadataService) => service.clearFreezePane('Sheet1'),
      },
      {
        bucket: 'sheetProtections',
        message: 'Failed to clear sheet protection metadata',
        run: (service: WorkbookMetadataService) => service.clearSheetProtection('Sheet1'),
      },
      {
        bucket: 'filters',
        message: 'Failed to delete filter metadata',
        run: (service: WorkbookMetadataService) => service.deleteFilter('Sheet1', range),
      },
      {
        bucket: 'sorts',
        message: 'Failed to delete sort metadata',
        run: (service: WorkbookMetadataService) => service.deleteSort('Sheet1', range),
      },
      {
        bucket: 'dataValidations',
        message: 'Failed to delete data validation metadata',
        run: (service: WorkbookMetadataService) => service.deleteDataValidation('Sheet1', range),
      },
      {
        bucket: 'conditionalFormats',
        message: 'Failed to delete conditional format metadata',
        run: (service: WorkbookMetadataService) => service.deleteConditionalFormat('cf-1'),
      },
      {
        bucket: 'rangeProtections',
        message: 'Failed to delete range protection metadata',
        run: (service: WorkbookMetadataService) => service.deleteRangeProtection('protect-1'),
      },
      {
        bucket: 'commentThreads',
        message: 'Failed to delete comment thread metadata',
        run: (service: WorkbookMetadataService) => service.deleteCommentThread('Sheet1', 'A1'),
      },
      {
        bucket: 'notes',
        message: 'Failed to delete note metadata',
        run: (service: WorkbookMetadataService) => service.deleteNote('Sheet1', 'A1'),
      },
      {
        bucket: 'spills',
        message: 'Failed to set spill metadata',
        run: (service: WorkbookMetadataService) => service.setSpill('Sheet1', 'A1', 2, 3),
      },
      {
        bucket: 'spills',
        message: 'Failed to delete spill metadata',
        run: (service: WorkbookMetadataService) => service.deleteSpill('Sheet1', 'A1'),
      },
      {
        bucket: 'spills',
        message: 'Failed to get spill metadata',
        run: (service: WorkbookMetadataService) => service.getSpill('Sheet1', 'A1'),
      },
      {
        bucket: 'spills',
        message: 'Failed to list spill metadata',
        run: (service: WorkbookMetadataService) => service.listSpills(),
      },
      {
        bucket: 'pivots',
        message: 'Failed to set pivot metadata',
        run: (service: WorkbookMetadataService) =>
          service.setPivot({
            name: 'RevenuePivot',
            sheetName: 'Sheet1',
            address: 'C3',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
            groupBy: ['Region'],
            values: [{ field: 'Sales', summarizeBy: 'sum' }],
            rows: 3,
            cols: 2,
          }),
      },
      {
        bucket: 'pivots',
        message: 'Failed to get pivot metadata',
        run: (service: WorkbookMetadataService) => service.getPivot('Sheet1', 'C3'),
      },
      {
        bucket: 'pivots',
        message: 'Failed to get pivot metadata by key',
        run: (service: WorkbookMetadataService) => service.getPivotByKey('Sheet1!C3'),
      },
      {
        bucket: 'pivots',
        message: 'Failed to delete pivot metadata',
        run: (service: WorkbookMetadataService) => service.deletePivot('Sheet1', 'C3'),
      },
      {
        bucket: 'pivots',
        message: 'Failed to list pivot metadata',
        run: (service: WorkbookMetadataService) => service.listPivots(),
      },
      {
        bucket: 'charts',
        message: 'Failed to set chart metadata',
        run: (service: WorkbookMetadataService) =>
          service.setChart({
            id: 'chart-1',
            sheetName: 'Sheet1',
            address: 'A1',
            source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
            chartType: 'line',
            rows: 2,
            cols: 2,
          }),
      },
      {
        bucket: 'charts',
        message: 'Failed to get chart metadata',
        run: (service: WorkbookMetadataService) => service.getChart('chart-1'),
      },
      {
        bucket: 'charts',
        message: 'Failed to delete chart metadata',
        run: (service: WorkbookMetadataService) => service.deleteChart('chart-1'),
      },
      {
        bucket: 'charts',
        message: 'Failed to list chart metadata',
        run: (service: WorkbookMetadataService) => service.listCharts(),
      },
      {
        bucket: 'images',
        message: 'Failed to set image metadata',
        run: (service: WorkbookMetadataService) =>
          service.setImage({
            id: 'image-1',
            sheetName: 'Sheet1',
            address: 'A1',
            sourceUrl: 'https://example.com/image.png',
            rows: 1,
            cols: 1,
          }),
      },
      {
        bucket: 'images',
        message: 'Failed to get image metadata',
        run: (service: WorkbookMetadataService) => service.getImage('image-1'),
      },
      {
        bucket: 'images',
        message: 'Failed to delete image metadata',
        run: (service: WorkbookMetadataService) => service.deleteImage('image-1'),
      },
      {
        bucket: 'images',
        message: 'Failed to list image metadata',
        run: (service: WorkbookMetadataService) => service.listImages(),
      },
      {
        bucket: 'shapes',
        message: 'Failed to set shape metadata',
        run: (service: WorkbookMetadataService) =>
          service.setShape({
            id: 'shape-1',
            sheetName: 'Sheet1',
            address: 'A1',
            shapeType: 'textBox',
            rows: 1,
            cols: 1,
          }),
      },
      {
        bucket: 'shapes',
        message: 'Failed to get shape metadata',
        run: (service: WorkbookMetadataService) => service.getShape('shape-1'),
      },
      {
        bucket: 'shapes',
        message: 'Failed to delete shape metadata',
        run: (service: WorkbookMetadataService) => service.deleteShape('shape-1'),
      },
      {
        bucket: 'shapes',
        message: 'Failed to list shape metadata',
        run: (service: WorkbookMetadataService) => service.listShapes(),
      },
    ] as const

    for (const testCase of cases) {
      const metadata = createWorkbookMetadataRecord() as Record<string, unknown>
      metadata[testCase.bucket] = createThrowingBucket(testCase.message)
      const service = createWorkbookMetadataService(metadata as ReturnType<typeof createWorkbookMetadataRecord>)
      expect(() => runWorkbookMetadataEffect(testCase.run(service))).toThrow(testCase.message)
    }
  })
})
