import type { WorkbookSnapshot, WorkPaperChange } from './work-paper-runtime-test-helpers.js'
import {
  cell,
  columnLabel,
  createWorkPaperFromDocument,
  describe,
  engineApplyCellMutationsTarget,
  engineExistingNumericCellMutationsTarget,
  engineFormulaBindingTarget,
  ErrorCode,
  expect,
  expectOnlyCellChanges,
  exportWorkPaperDocument,
  hasCaptureVisibilitySnapshot,
  hasDeferredTrackedIndexChanges,
  it,
  parseWorkPaperDocument,
  readUndoStack,
  rejectSingleTrackedCellReader,
  serializeWorkPaperDocument,
  sheetGridEntryTarget,
  trackCaptureTrackedChangesWithoutVisibilityCache,
  trackComputeCellChangesFromTrackedEvents,
  trackPrivateMethod,
  trackSheetDimensionCacheUpdates,
  ValueTag,
  vi,
  WorkPaper,
} from './work-paper-runtime-test-helpers.js'

describe('WorkPaper lifecycle, snapshots, and tracked mutation payloads', () => {
  it('builds from named sheets and exposes stable sheet ids and serialization helpers', () => {
    const workbook = WorkPaper.buildFromSheets({
      Summary: [[1, '=A1*2']],
      Detail: [[3]],
    })

    const summaryId = workbook.getSheetId('Summary')!

    expect(workbook.getSheetName(summaryId)).toBe('Summary')
    expect(workbook.countSheets()).toBe(2)
    expect(workbook.getCellValue(cell(summaryId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(workbook.getCellFormula(cell(summaryId, 0, 1))).toBe('=A1*2')
    expect(workbook.getCellSerialized(cell(summaryId, 0, 1))).toBe('=A1*2')
    expect(workbook.getSheetDimensions(summaryId)).toEqual({ width: 2, height: 1 })
    expect(workbook.simpleCellAddressFromString('Summary!B1')).toEqual(cell(summaryId, 0, 1))
    expect(workbook.simpleCellRangeFromString('Summary!A1:B1')).toEqual({
      start: cell(summaryId, 0, 0),
      end: cell(summaryId, 0, 1),
    })
  })

  it('rejects public metadata and dependency reads after disposal', () => {
    const workbook = WorkPaper.buildFromSheets({
      Summary: [[1, '=A1*2']],
    })
    const summaryId = workbook.getSheetId('Summary')!

    workbook.dispose()

    const disposedReads: readonly (() => unknown)[] = [
      () => workbook.getConfig(),
      () => workbook.getSheetName(summaryId),
      () => workbook.getSheetNames(),
      () => workbook.getSheetId('Summary'),
      () => workbook.doesSheetExist('Summary'),
      () => workbook.countSheets(),
      () => workbook.isThereSomethingToUndo(),
      () => workbook.isThereSomethingToRedo(),
      () => workbook.clearUndoStack(),
      () => workbook.clearRedoStack(),
      () => workbook.getCellPrecedents(cell(summaryId, 0, 1)),
      () => workbook.getCellDependents(cell(summaryId, 0, 0)),
      () => workbook.isItPossibleToSetCellContents(cell(summaryId, 0, 0), 2),
      () => workbook.isItPossibleToAddRows(summaryId, 1, 1),
      () => workbook.isItPossibleToMoveRows(summaryId, 0, 1, 1),
      () => workbook.isItPossibleToAddSheet('Next'),
      () => workbook.isItPossibleToRemoveSheet(summaryId),
      () => workbook.isItPossibleToAddNamedExpression('DisposedName', '=1'),
      () => workbook.isItPossibleToRemoveNamedExpression('DisposedName'),
      () => workbook.licenseKeyValidityState,
      () => workbook.graph,
      () => workbook.sheetMapping,
      () => workbook.dependencyGraph,
    ]

    disposedReads.forEach((read) => {
      expect(read).toThrow('Workbook has been disposed')
    })
  })

  it('does not leak conditional aggregation caches across disposed workbooks', () => {
    const first = WorkPaper.buildFromSheets({
      Bench: [
        ['Group', 'Value', '', 'A', '=SUMIF(A2:A5,D1,B2:B5)'],
        ['A', 1],
        ['B', 2],
        ['A', 3],
        ['B', 4],
      ],
    })
    const firstSheetId = first.getSheetId('Bench')!

    expect(first.getCellValue(cell(firstSheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    })
    first.dispose()

    const second = WorkPaper.buildFromSheets({
      Bench: [
        ['Group', 'Value', '', 'A', '=SUMIF(A2:A5,D1,B2:B5)'],
        ['B', 10],
        ['B', 20],
        ['B', 30],
        ['B', 40],
      ],
    })
    const secondSheetId = second.getSheetId('Bench')!

    expect(second.getCellValue(cell(secondSheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 0,
    })
    second.dispose()
  })

  it('does not reuse stale sheet names for pooled formula initialization', () => {
    const first = WorkPaper.buildFromSheets({
      Cases: [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9, '=SUM(INDEX(A1:C3,0,2))'],
      ],
    })
    const firstSheetId = first.getSheetId('Cases')!

    expect(first.getCellValue(cell(firstSheetId, 2, 3))).toEqual({
      tag: ValueTag.Number,
      value: 15,
    })
    first.dispose()

    const second = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, 2, 3, '', '', '', '=SUM(INDEX(A1:C3,0,2))', '=SUM(INDEX(A1:C3,2,0))', '=SUM(INDEX(A1:C3,0,0))'],
        [4, 5, 6],
        [7, 8, 9],
      ],
    })
    const secondSheetId = second.getSheetId('Sheet1')!

    expect([6, 7, 8].map((col) => second.getCellValue(cell(secondSheetId, 0, col)))).toEqual([
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.Number, value: 45 },
    ])
    second.dispose()
  })

  it('does not reuse stale region subscriptions for pooled spill recalculation', () => {
    const first = WorkPaper.buildFromSheets({
      Cases: [
        [
          'North',
          10,
          '=COUNTBLANK(A1:A5)',
          '=COUNTIF(A1:A5,"")',
          '=COUNTIF(A1:A5,"<>")',
          '=SUMIF(A1:A5,"",B1:B5)',
          '=SUMIF(A1:A5,"<>",B1:B5)',
          '=SUMIFS(B1:B5,A1:A5,"<>")',
        ],
        [null, 20],
        ['=IF(TRUE,"","x")', 30],
        [' ', 40],
        ['South', 50],
      ],
    })
    first.dispose()

    const second = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, 2, 3, null, '=OFFSET(A1,0,1,3,1)', null, '=SUM(OFFSET(A1,0,1,3,1))'],
        [4, 5, 6],
        [7, 8, 9],
      ],
    })
    second.dispose()

    const third = WorkPaper.buildFromSheets({
      Cases: [[3, '=SEQUENCE(A1,1,1,1)', null, '=SUM(B1#)', '=ROWS(B1#)', '=IFERROR(INDEX(B1#,2),"missing")'], [], []],
    })
    const sheetId = third.getSheetId('Cases')!
    third.setCellContents(cell(sheetId, 0, 0), 1)

    expect([1, 2, 3, 4, 5].map((col) => third.getCellValue(cell(sheetId, 0, col)))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'missing', stringId: expect.any(Number) },
    ])
    expect([1, 2].map((row) => third.getCellValue(cell(sheetId, row, 1)))).toEqual([{ tag: ValueTag.Empty }, { tag: ValueTag.Empty }])
    third.dispose()
  })

  it('builds from imported workbook snapshots with metadata-backed formulas', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Structured Financial Model',
        metadata: {
          definedNames: [
            { name: 'Currency', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'F7' } },
            { name: 'Start_Year', value: { kind: 'cell-ref', sheetName: 'Constants', address: 'B10' } },
          ],
          tables: [
            {
              name: 'tblActuals',
              sheetName: 'Imports',
              startAddress: 'A6',
              endAddress: 'D8',
              columnNames: ['Account', 'Value', 'Year', 'Period'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Constants',
          order: 0,
          cells: [
            { address: 'B10', value: 2012 },
            { address: 'F7', value: 'USD' },
            { address: 'F9', formula: 'Currency & "  000s"' },
          ],
        },
        {
          id: 2,
          name: 'Imports',
          order: 1,
          cells: [
            { address: 'A6', value: 'Account' },
            { address: 'B6', value: 'Value' },
            { address: 'C6', value: 'Year' },
            { address: 'D6', value: 'Period' },
            { address: 'A7', value: 'Revenue' },
            { address: 'B7', value: 100 },
            { address: 'C7', value: 2011 },
            { address: 'D7', formula: "'Imports'!C7-Start_Year+1" },
            { address: 'A8', value: 'Revenue' },
            { address: 'B8', value: 125 },
            { address: 'C8', value: 2012 },
            { address: 'D8', formula: "'Imports'!C8-Start_Year+1" },
            { address: 'F10', formula: "SUM('Imports'!B7:B8)" },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 20, maxColumns: 8, useColumnIndex: true })
    const constantsId = workbook.getSheetId('Constants')!
    const importsId = workbook.getSheetId('Imports')!

    expect(workbook.getCellValue(cell(constantsId, 8, 5))).toMatchObject({ tag: ValueTag.String, value: 'USD  000s' })
    expect(workbook.getCellValue(cell(importsId, 6, 3))).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(workbook.getCellValue(cell(importsId, 7, 3))).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(workbook.getCellValue(cell(importsId, 9, 5))).toEqual({ tag: ValueTag.Number, value: 225 })
    expect(workbook.getSheetDimensions(importsId)).toEqual({ width: 6, height: 10 })

    workbook.setCellContents(cell(importsId, 6, 2), 2013)

    expect(workbook.getCellValue(cell(importsId, 6, 3))).toEqual({ tag: ValueTag.Number, value: 2 })

    const exported = workbook.exportSnapshot()
    expect(exported.workbook.metadata?.definedNames).toEqual(snapshot.workbook.metadata?.definedNames)
    expect(exported.workbook.metadata?.tables).toEqual(snapshot.workbook.metadata?.tables)
    expect(exported.sheets.find((sheet) => sheet.name === 'Imports')?.cells).toContainEqual({ address: 'C7', value: 2013 })
    workbook.dispose()
  })

  it('keeps literal-only initialization compatible with named expressions and later formulas', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          [2, 'west', true],
          [4, null, false],
        ],
      },
      {},
      [{ name: 'BenchTotal', expression: '=SUM(Bench!$A$1:$A$2)' }],
    )
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getNamedExpressionValue('BenchTotal')).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })

    const changes = workbook.setCellContents(cell(sheetId, 0, 3), '=BenchTotal+A1')

    expect(changes).toHaveLength(1)
    expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(workbook.getSheetSerialized(sheetId)).toEqual([
      [2, 'west', true, '=BenchTotal+A1'],
      [4, null, false, null],
    ])
  })

  it('builds mixed literal and formula sheets without seeding undo history', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 10, 'label-1', true, '=A1+B1', '=E1*2'],
        [2, 20, 'label-2', false, '=A2+B2', '=E2*2'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 5))).toEqual({
      tag: ValueTag.Number,
      value: 44,
    })
    expect(workbook.isThereSomethingToUndo()).toBe(false)

    const changes = workbook.setCellContents(cell(sheetId, 1, 1), 30)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
      'Bench!B2',
      'Bench!E2',
      'Bench!F2',
    ])
    expect(workbook.getCellValue(cell(sheetId, 1, 5))).toEqual({
      tag: ValueTag.Number,
      value: 64,
    })
  })

  it('uses engine-emitted changed-cell payloads for ordinary value edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, '=A1*2']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('ordinary value edits should not rebuild visibility snapshots')
    })

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 7)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 14,
    })
    captureVisibilitySnapshot.mockRestore()
  })

  it('uses direct changed-cell payloads for existing string leaf formula edits', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [['foo', '=CONCATENATE(A1,"-bar")']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('string leaf edits should not rebuild visibility snapshots')
    })

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 'baz')

    expect(changes).toEqual([
      {
        kind: 'cell',
        address: cell(sheetId, 0, 0),
        sheetName: 'Bench',
        a1: 'A1',
        newValue: expect.objectContaining({ tag: ValueTag.String, value: 'baz' }),
      },
      {
        kind: 'cell',
        address: cell(sheetId, 0, 1),
        sheetName: 'Bench',
        a1: 'B1',
        newValue: expect.objectContaining({ tag: ValueTag.String, value: 'baz-bar' }),
      },
    ])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({ tag: ValueTag.String, value: 'baz-bar' })
    expect(workbook.getPerformanceCounters().changedCellPayloadsBuilt).toBe(0)
    expect(workbook.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
    captureVisibilitySnapshot.mockRestore()
  })

  it('uses tracked patch payloads without exposing internal cell indices', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, '=A1*2']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
    expect(changes.every((change) => change.kind !== 'cell' || !('cellIndex' in change))).toBe(true)
    expect(workbook.getPerformanceCounters().changedCellPayloadsBuilt).toBe(0)
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 18,
    })
  })

  it('uses a direct tracked payload for single literal edits without core materialization', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1]],
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toEqual([
      {
        kind: 'cell',
        address: cell(sheetId, 0, 0),
        sheetName: 'Bench',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 9 },
      },
    ])
    expect(workbook.getPerformanceCounters().changedCellPayloadsBuilt).toBe(0)
  })

  it('returns large physical formula replacement changes lazily without the generic tracked reducer', () => {
    const downstreamCount = 64
    const row: unknown[] = [1, 2, '=A1+B1']
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      row.push(`=${columnLabel(col - 1)}1+1`)
    }
    const workbook = WorkPaper.buildFromSheets({ Bench: [row] })
    const sheetId = workbook.getSheetId('Bench')!
    const trackedReducer = trackComputeCellChangesFromTrackedEvents(workbook)

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.setCellContents(cell(sheetId, 0, 2), '=A1*B1')

      expect(changes).toHaveLength(downstreamCount + 1)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(trackedReducer.count).toBe(0)
      expect(workbook.getPerformanceCounters().changedCellPayloadsBuilt).toBe(0)
      expect(workbook.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
      expect(workbook.getCellValue(cell(sheetId, 0, downstreamCount + 2))).toEqual({
        tag: ValueTag.Number,
        value: downstreamCount + 2,
      })
    } finally {
      trackedReducer.restore()
    }
  })

  it('keeps formula-family indexing lazy when undo-capturing ordinary formula replacements', () => {
    const downstreamCount = 64
    const row: unknown[] = [1, 2, '=A1+B1']
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = 2 + offset
      row.push(`=${columnLabel(col - 1)}1+1`)
    }
    const workbook = WorkPaper.buildFromSheets({ Bench: [row] })
    const sheetId = workbook.getSheetId('Bench')!
    const binding = engineFormulaBindingTarget(workbook)

    expect(binding.isFormulaFamilyIndexReadyNow()).toBe(false)

    workbook.setCellContents(cell(sheetId, 0, 2), '=A1*B1')

    expect(binding.isFormulaFamilyIndexReadyNow()).toBe(false)
    expect(workbook.getCellFormula(cell(sheetId, 0, 2))).toBe('=A1*B1')
    expect(workbook.getCellValue(cell(sheetId, 0, downstreamCount + 2))).toEqual({
      tag: ValueTag.Number,
      value: downstreamCount + 2,
    })
    expect(workbook.undo()).toHaveLength(downstreamCount + 1)
    expect(workbook.getCellFormula(cell(sheetId, 0, 2))).toBe('=A1+B1')
  })

  it('returns tiny no-listener compact tracked changes eagerly', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, '=A1*2']],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toHaveLength(2)
    expectOnlyCellChanges(changes)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
  })

  it('keeps valuesUpdated listener payloads eager for tiny tracked changes', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, '=A1*2']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    const events: WorkPaperChange[][] = []
    workbook.on('valuesUpdated', (changes) => {
      events.push(changes)
    })

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(events).toHaveLength(1)
    expect(events[0]).toBe(changes)
    expectOnlyCellChanges(changes)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)
    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
  })

  it('updates small sliding aggregate fanout without dirty traversal', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: 64 }, (_, row) => {
        const rowNumber = row + 1
        const endRow = Math.min(64, rowNumber + 31)
        return [rowNumber, `=SUM(A${rowNumber}:A${endRow})`]
      }),
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 10)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 537,
    })
    expect(workbook.getStats().lastMetrics).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directAggregateDeltaApplications: 1,
      directAggregateDeltaOnlyRecalcSkips: 1,
      regionQueryIndexBuilds: 0,
    })
  })

  it('recalculates filter spills that share a dirty range with direct criteria formulas', () => {
    const workbook = WorkPaper.buildFromSheets({
      Deals: [
        ['Region', 'Segment', 'Customers', 'ARPA', 'Revenue'],
        ['West', 'Enterprise', 12, 1200, '=C2*D2'],
        ['East', 'SMB', 30, 250, '=C3*D3'],
        ['West', 'SMB', 18, 300, '=C4*D4'],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Total revenue', '=SUM(Deals!E2:E4)'],
        ['West customers', '=SUMIF(Deals!A2:A4,"West",Deals!C2:C4)'],
        ['Qualified customer counts', '=FILTER(Deals!C2:C4,Deals!C2:C4>=18)'],
      ],
    })
    const dealsSheet = workbook.getSheetId('Deals')!
    const summarySheet = workbook.getSheetId('Summary')!
    const readQualifiedCounts = (target: WorkPaper, sheet: number) =>
      target
        .getRangeValues({
          start: cell(sheet, 3, 1),
          end: cell(sheet, 5, 1),
        })
        .flat()
        .map((value) => (value.tag === ValueTag.Number ? value.value : null))

    expect(readQualifiedCounts(workbook, summarySheet)).toEqual([30, 18, null])

    workbook.setCellContents(cell(dealsSheet, 1, 2), 20)

    expect(readQualifiedCounts(workbook, summarySheet)).toEqual([20, 30, 18])

    const restored = createWorkPaperFromDocument(
      parseWorkPaperDocument(serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))),
    )

    expect(readQualifiedCounts(restored, restored.getSheetId('Summary')!)).toEqual([20, 30, 18])
  })

  it('captures tiny sliding aggregate listener payloads without the general tracked reducer', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: 64 }, (_, row) => {
        const rowNumber = row + 1
        const endRow = Math.min(64, rowNumber + 31)
        return [rowNumber, `=SUM(A${rowNumber}:A${endRow})`]
      }),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const reducerTracker = trackComputeCellChangesFromTrackedEvents(workbook)
    const captureTracker = trackCaptureTrackedChangesWithoutVisibilityCache(workbook)
    const genericReader = rejectSingleTrackedCellReader(workbook)
    workbook.on('valuesUpdated', () => {})

    try {
      const changes = workbook.setCellContents(cell(sheetId, 0, 0), 10)

      expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!A1', 'Bench!B1'])
      expect(captureTracker.count).toBe(0)
      expect(reducerTracker.count).toBe(0)
    } finally {
      genericReader.restore()
      captureTracker.restore()
      reducerTracker.restore()
    }
  })

  it('captures tiny indexed lookup listener payloads without the general tracked reducer', () => {
    const rowCount = 64
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', 32, `=MATCH(D1,A2:A${rowCount + 1},0)`],
          ...Array.from({ length: rowCount }, (_, row) => [row + 1, (row + 1) * 10]),
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!
    const reducerTracker = trackComputeCellChangesFromTrackedEvents(workbook)
    workbook.on('valuesUpdated', () => {})

    try {
      const changes = workbook.setCellContents(cell(sheetId, rowCount, 0), rowCount + 1_000)

      expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
        `Bench!A${rowCount + 1}`,
      ])
      expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
        tag: ValueTag.Number,
        value: 32,
      })
      expect(reducerTracker.count).toBe(0)
    } finally {
      reducerTracker.restore()
    }
  })

  it('updates indexed exact text lookup operands without dirty traversal or index rebuilds', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', 'alpha', '=MATCH(D1,A2:A5,0)'],
          ['alpha', 10],
          ['bravo', 20],
          ['charlie', 30],
          ['delta', 40],
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, 0, 3), 'delta')

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!D1', 'Bench!E1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(workbook.getStats().lastMetrics).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      exactIndexBuilds: 0,
      lookupOwnerBuilds: 0,
    })
  })

  it('updates non-uniform approximate lookup operands through prepared numeric vectors', () => {
    const rowCount = 64
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', 20, `=MATCH(D1,A2:A${rowCount + 1},1)`],
        ...Array.from({ length: rowCount }, (_, row) => [Math.ceil((row + 1) / 2), (row + 1) * 10]),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, 0, 3), 11)

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Bench!D1', 'Bench!E1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(workbook.getStats().lastMetrics).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      approxIndexBuilds: 0,
    })
  })

  it('uses bulk tracked indices for large literal batches without core patch payloads', () => {
    const rowCount = 600
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: rowCount }, (_, row) => [row + 1, `=A${row + 1}*2`]),
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.batch(() => {
      for (let row = 0; row < rowCount; row += 1) {
        workbook.setCellContents(cell(sheetId, row, 0), row * 3)
      }
    })

    expect(changes).toHaveLength(rowCount * 2)
    expect(changes.every((change) => change.kind !== 'cell' || !('cellIndex' in change))).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 1))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 6,
    })
    expect(workbook.getPerformanceCounters().changedCellPayloadsBuilt).toBe(0)
  })

  it('keeps large multi-column batches on the deferred physical tracked path without visibility snapshots', () => {
    const rowCount = 128
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: rowCount }, (_, row) => {
        const rowNumber = row + 1
        return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=A${rowNumber}*B${rowNumber}`]
      }),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('large no-listener physical batches should not rebuild visibility snapshots')
    })
    const dimensionUpdates = trackSheetDimensionCacheUpdates(workbook)
    const genericReader = rejectSingleTrackedCellReader(workbook)

    try {
      const changes = workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(cell(sheetId, row, 0), row * 3)
          workbook.setCellContents(cell(sheetId, row, 1), row * 5)
        }
      })

      expect(changes).toHaveLength(rowCount * 4)
      expectOnlyCellChanges(changes)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 8,
      })
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 3))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
      })
      expect(dimensionUpdates.count).toBe(0)
    } finally {
      genericReader.restore()
      dimensionUpdates.restore()
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('reads physical dense ranges without entering the core read service', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2],
        ['x', true],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    const engine = Reflect.get(workbook, 'engine')
    const getRangeValues = vi.spyOn(engine, 'getRangeValues').mockImplementation(() => {
      throw new Error('physical range reads should use the headless fast path')
    })

    const values = workbook.getRangeValues({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 1, 1),
    })

    expect(values).toEqual([
      [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ],
      [
        { tag: ValueTag.String, value: 'x', stringId: expect.any(Number) },
        { tag: ValueTag.Boolean, value: true },
      ],
    ])
    expect(getRangeValues).not.toHaveBeenCalled()
    getRangeValues.mockRestore()
  })

  it('keeps large physical dense range reads on the headless fast path', () => {
    const rowCount = 260
    const colCount = 64
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: rowCount }, (_rowValue, row) =>
        Array.from({ length: colCount }, (_colValue, col) => row * colCount + col + 1),
      ),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const engine = Reflect.get(workbook, 'engine')
    const getRangeValues = vi.spyOn(engine, 'getRangeValues').mockImplementation(() => {
      throw new Error('large physical range reads should use the headless fast path')
    })

    const values = workbook.getRangeValues({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, rowCount - 1, colCount - 1),
    })

    expect(values).toHaveLength(rowCount)
    expect(values[0]).toHaveLength(colCount)
    expect(values[0]?.[0]).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(values[rowCount - 1]?.[colCount - 1]).toEqual({
      tag: ValueTag.Number,
      value: rowCount * colCount,
    })
    expect(getRangeValues).not.toHaveBeenCalled()
    getRangeValues.mockRestore()
  })

  it('keeps structurally edited dense range reads on the headless fast path', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, 3],
        [4, 5, 6],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.removeColumns(sheetId, [1, 1])
    const engine = Reflect.get(workbook, 'engine')
    const getRangeValues = vi.spyOn(engine, 'getRangeValues').mockImplementation(() => {
      throw new Error('logical range reads should use the headless fast path')
    })

    const values = workbook.getRangeValues({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 1, 1),
    })

    expect(values).toEqual([
      [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ],
      [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ],
    ])
    expect(getRangeValues).not.toHaveBeenCalled()
    getRangeValues.mockRestore()
  })

  it('uses initialized sheet dimensions without scanning existing-cell batch grids', () => {
    const workbook = WorkPaper.buildFromArray([
      [1, 2],
      [3, 4],
    ])
    const sheetId = workbook.getSheetId('Sheet1')!
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 2 })
      expect(forEachCellEntry).not.toHaveBeenCalled()

      workbook.batch(() => {
        workbook.setCellContents(cell(sheetId, 0, 0), 10)
        workbook.setCellContents(cell(sheetId, 1, 1), 40)
      })

      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 2 })
      expect(forEachCellEntry).not.toHaveBeenCalled()

      workbook.setCellContents(cell(sheetId, 3, 4), 99)
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 5, height: 4 })
      expect(forEachCellEntry).not.toHaveBeenCalled()

      workbook.setCellContents(cell(sheetId, 3, 4), null)
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 5, height: 4 })
      expect(forEachCellEntry).toHaveBeenCalledTimes(1)
    } finally {
      forEachCellEntry.mockRestore()
    }
  })

  it('supports sheet-scoped named expressions and restores public formulas', () => {
    const workbook = WorkPaper.buildFromSheets({
      Summary: [[]],
      Detail: [[]],
    })
    const summaryId = workbook.getSheetId('Summary')!
    const detailId = workbook.getSheetId('Detail')!
    const events: string[] = []

    workbook.on('namedExpressionAdded', (name, changes) => {
      events.push(`add:${name}:${changes.length}`)
    })
    workbook.onDetailed('namedExpressionAdded', (payload) => {
      events.push(`scope:${payload.scope}`)
    })
    workbook.on('valuesUpdated', (changes) => {
      events.push(`values:${changes.length}`)
    })

    workbook.addNamedExpression('Rate', '=1', summaryId)
    workbook.addNamedExpression('Rate', '=2', detailId)

    expect(workbook.setCellContents(cell(summaryId, 0, 0), '=Rate+1')).toHaveLength(1)
    expect(workbook.setCellContents(cell(detailId, 0, 0), '=Rate+1')).toHaveLength(1)

    expect(workbook.getCellValue(cell(summaryId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(workbook.getCellValue(cell(detailId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(workbook.getCellFormula(cell(summaryId, 0, 0))).toBe('=Rate+1')
    expect(workbook.getCellFormula(cell(detailId, 0, 0))).toBe('=Rate+1')
    expect(workbook.getNamedExpressionValue('Rate', summaryId)).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(workbook.getNamedExpressionValue('Rate', detailId)).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(events.slice(0, 6)).toEqual(['add:Rate:1', 'scope:1', 'values:1', 'add:Rate:1', 'scope:2', 'values:1'])
  })

  it('coalesces batch history into one undo entry and emits one values update', () => {
    const workbook = WorkPaper.buildFromArray([[1]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const valuesUpdated: number[] = []
    const nestedMutationResults: number[] = []

    workbook.on('valuesUpdated', (changes) => {
      valuesUpdated.push(changes.length)
    })

    const changes = workbook.batch(() => {
      nestedMutationResults.push(workbook.setCellContents(cell(sheetId, 0, 1), '=A1*2').length)
      nestedMutationResults.push(workbook.setCellContents(cell(sheetId, 1, 0), 5).length)
    })

    expect(changes).toHaveLength(2)
    expect(nestedMutationResults).toEqual([0, 0])
    expect(valuesUpdated).toEqual([2])
    expect(workbook.isThereSomethingToUndo()).toBe(true)

    const undoChanges = workbook.undo()

    expect(undoChanges).toHaveLength(2)
    expect(workbook.getCellValue(cell(sheetId, 0, 1)).tag).toBe(ValueTag.Empty)
    expect(workbook.getCellValue(cell(sheetId, 1, 0)).tag).toBe(ValueTag.Empty)
  })

  it('uses tracked engine changes for literal-only outer batches on a fresh workbook', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('literal-only outer batches should not rebuild visibility snapshots')
    })

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.setCellContents(cell(sheetId, 1, 0), 20)
    })

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Sheet1!A1', 'Sheet1!A2'])
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    })
    captureVisibilitySnapshot.mockRestore()
  })

  it('coalesces repeated tracked batch writes to the same cell', () => {
    const workbook = WorkPaper.buildFromArray([[1, '=A1*2']])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 2)
      workbook.setCellContents(cell(sheetId, 0, 0), 3)
    })

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual(['Sheet1!A1', 'Sheet1!B1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 6 })
  })

  it('keeps merged existing-numeric batch history on typed mutation records', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!

    workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.setCellContents(cell(sheetId, 1, 0), 20)
    })

    const undoStack = readUndoStack(workbook)
    expect(undoStack).not.toBeNull()
    expect(undoStack).toHaveLength(1)
    expect(Reflect.get(undoStack?.[0], 'forward') ? Reflect.get(Reflect.get(undoStack?.[0], 'forward'), 'kind') : undefined).toBe(
      'existing-numeric-cell-mutations',
    )
    expect(Reflect.get(undoStack?.[0], 'inverse') ? Reflect.get(Reflect.get(undoStack?.[0], 'inverse'), 'kind') : undefined).toBe(
      'existing-numeric-cell-mutations',
    )
  })

  it('passes known zero potential-new-cell count for existing-cell batch flushes', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const applyCellMutationsAt = vi.spyOn(engineApplyCellMutationsTarget(workbook), 'applyCellMutationsAtWithOptions')
    const applyExistingNumericBatch = vi.spyOn(engineExistingNumericCellMutationsTarget(workbook), 'tryApplyExistingNumericCellMutationsAt')

    try {
      workbook.batch(() => {
        workbook.setCellContents(cell(sheetId, 0, 0), 10)
        workbook.setCellContents(cell(sheetId, 1, 0), 20)
      })

      expect(applyCellMutationsAt).not.toHaveBeenCalled()
      expect(applyExistingNumericBatch).toHaveBeenCalledTimes(1)
      const request = Object(applyExistingNumericBatch.mock.calls[0]?.[0])
      expect(Reflect.get(request, 'potentialNewCells')).toBe(0)
      expect(Reflect.get(request, 'sheetIds')).toHaveLength(2)
    } finally {
      applyCellMutationsAt.mockRestore()
      applyExistingNumericBatch.mockRestore()
    }
  })

  it('passes known zero potential-new-cell count for existing-cell suspended flushes', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const dimensionUpdates = trackPrivateMethod(workbook, 'updateSheetDimensionsAfterCellMutationRefs')

    workbook.suspendEvaluation()
    const applyCellMutationsAt = vi.spyOn(engineApplyCellMutationsTarget(workbook), 'applyCellMutationsAtWithOptions')

    try {
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.setCellContents(cell(sheetId, 1, 0), 20)
      expect(applyCellMutationsAt).not.toHaveBeenCalled()

      workbook.resumeEvaluation()

      expect(applyCellMutationsAt).toHaveBeenCalledTimes(1)
      expect(applyCellMutationsAt.mock.calls[0]?.[1]).toMatchObject({
        captureUndo: true,
        potentialNewCells: 0,
        reuseRefs: true,
        source: 'local',
      })
      expect(dimensionUpdates.count).toBe(0)
    } finally {
      applyCellMutationsAt.mockRestore()
      dimensionUpdates.restore()
    }
  })

  it('flushes deferred literal edits before formula writes inside a batch', () => {
    const workbook = WorkPaper.buildFromArray([[1]])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.batch(() => {
      expect(workbook.setCellContents(cell(sheetId, 0, 0), 10)).toEqual([])
      expect(workbook.setCellContents(cell(sheetId, 0, 1), '=A1*2')).toEqual([])
    })

    expect(changes).toHaveLength(2)
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    })
  })

  it('undoes and redoes deferred literal-only batches', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.setCellContents(cell(sheetId, 1, 0), 20)
    })

    expect(changes).toHaveLength(2)
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    })

    workbook.undo()
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })

    workbook.redo()
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    })
  })

  it('returns stable array-compatible tracked changes for large direct batches', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: 32 }, (_, index) => [index + 1, `=A${index + 1}*2`]),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const emittedChanges: WorkPaperChange[][] = []

    workbook.on('valuesUpdated', (changes) => {
      emittedChanges.push(changes)
    })

    const changes = workbook.batch(() => {
      for (let row = 0; row < 32; row += 1) {
        workbook.setCellContents(cell(sheetId, row, 0), row * 3)
      }
    })

    expect(Array.isArray(changes)).toBe(true)
    expect(changes).toHaveLength(64)
    expect(emittedChanges).toHaveLength(1)
    expect(emittedChanges[0]).toBe(changes)
    expect(changes.slice(0, 4).map((change) => (change.kind === 'cell' ? change.a1 : change.kind))).toEqual(['A1', 'B1', 'A2', 'B2'])
    const firstChange = changes[0]
    expect(firstChange).toMatchObject({
      kind: 'cell',
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 0 },
    })

    workbook.setCellContents(cell(sheetId, 0, 0), 999)

    expect(changes[0]).toBe(firstChange)
    expect(firstChange).toMatchObject({
      kind: 'cell',
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 0 },
    })
    expect(changes[1]).toMatchObject({
      kind: 'cell',
      a1: 'B1',
      newValue: { tag: ValueTag.Number, value: 0 },
    })
    expect(changes[63]).toMatchObject({
      kind: 'cell',
      a1: 'B32',
      newValue: { tag: ValueTag.Number, value: 186 },
    })
  })

  it('returns stable array-compatible tracked changes for large suspended batches', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: 32 }, (_, index) => [index + 1, `=A${index + 1}*2`]),
    })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.suspendEvaluation()
    for (let row = 0; row < 32; row += 1) {
      workbook.setCellContents(cell(sheetId, row, 0), row * 7)
    }
    const changes = workbook.resumeEvaluation()

    expect(Array.isArray(changes)).toBe(true)
    expect(changes).toHaveLength(64)
    const firstChange = changes[0]
    expect(firstChange).toMatchObject({
      kind: 'cell',
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 0 },
    })

    workbook.setCellContents(cell(sheetId, 0, 0), 999)

    expect(changes[0]).toBe(firstChange)
    expect(changes[1]).toMatchObject({
      kind: 'cell',
      a1: 'B1',
      newValue: { tag: ValueTag.Number, value: 0 },
    })
    expect(changes[63]).toMatchObject({
      kind: 'cell',
      a1: 'B32',
      newValue: { tag: ValueTag.Number, value: 434 },
    })
  })

  it('keeps exact MATCH correct when useColumnIndex is enabled', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [[1, '', '', 2, '=MATCH(D1,A1:A3,0)'], [2], [3]],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })

    const missingMatchChanges = workbook.setCellContents(cell(sheetId, 1, 0), 20)
    expect(missingMatchChanges.map((change) => (change.kind === 'cell' ? change.a1 : change.kind))).toEqual(['E1', 'A2'])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    const restoredMatchChanges = workbook.setCellContents(cell(sheetId, 0, 3), 3)
    expect(restoredMatchChanges.map((change) => (change.kind === 'cell' ? change.a1 : change.kind))).toEqual(['D1', 'E1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 3,
    })
  })

  it('applies duplicate approximate MATCH operand edits through the compact direct path', () => {
    const rowCount = 64
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', Math.floor(rowCount / 4), `=MATCH(D1,A2:A${rowCount + 1},1)`],
        ...Array.from({ length: rowCount }, (_, index) => {
          const key = Math.ceil((index + 1) / 2)
          return [key, key * 10]
        }),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: rowCount / 2 })
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, 0, 3), rowCount / 4 + 4)

    expect(changes.map((change) => (change.kind === 'cell' ? change.a1 : change.kind))).toEqual(['D1', 'E1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: rowCount / 2 + 8 })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      changedCellPayloadsBuilt: 0,
      lookupOwnerBuilds: 0,
    })
  })

  it('defers kernel sync for lookup-column writes with no dirty formula dependents', () => {
    const rowCount = 64
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        ['Key', 'Value', '', Math.floor(rowCount / 2), `=MATCH(D1,A2:A${rowCount + 1},1)`],
        ...Array.from({ length: rowCount }, (_, index) => [index + 1, (index + 1) * 10]),
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    const changes = workbook.setCellContents(cell(sheetId, rowCount, 0), rowCount + 1)

    expect(changes.map((change) => (change.kind === 'cell' ? change.a1 : change.kind))).toEqual([`A${rowCount + 1}`])
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(rowCount / 2),
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      kernelSyncOnlyRecalcSkips: 1,
      wasmFullUploads: 0,
    })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    })
  })
})
