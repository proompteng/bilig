import type { WorkbookSnapshot } from './work-paper-runtime-test-helpers.js'
import {
  cell,
  describe,
  engineApplyCellMutationsTarget,
  engineFormulaBindingTarget,
  ErrorCode,
  expect,
  exportWorkPaperDocument,
  hasCaptureVisibilitySnapshot,
  it,
  readEngineUseColumnIndexEnabled,
  readFormulaBindingTestSurface,
  sheetGridEntryTarget,
  TEST_LANGUAGE_CODE,
  trackCaptureTrackedChangesWithoutVisibilityCache,
  trackComputeCellChangesFromTrackedEvents,
  trackPrivateMethod,
  ValueTag,
  vi,
  WorkPaper,
  WorkPaperEvaluationSuspendedError,
} from './work-paper-runtime-test-helpers.js'

describe('WorkPaper cache, suspended evaluation, and formula edge cases', () => {
  it('applies suspended exact lookup-column writes as a tracked input-only batch', () => {
    const rowCount = 96
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [
          ['Key', 'Value', '', Math.floor(rowCount / 2), `=MATCH(D1,A2:A${rowCount + 1},0)`],
          ...Array.from({ length: rowCount }, (_, index) => [index + 1, (index + 1) * 10]),
        ],
      },
      { useColumnIndex: true },
    )
    const sheetId = workbook.getSheetId('Bench')!
    workbook.resetPerformanceCounters()

    workbook.suspendEvaluation()
    for (let index = 0; index < 32; index += 1) {
      const row = rowCount - index
      workbook.setCellContents(cell(sheetId, row, 0), row + 1_000)
    }
    const changes = workbook.resumeEvaluation()

    expect(changes).toHaveLength(32)
    expect(changes[0]).toMatchObject({ kind: 'cell', a1: `A${rowCount - 30}` })
    expect(changes.at(-1)).toMatchObject({ kind: 'cell', a1: `A${rowCount + 1}` })
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(rowCount / 2),
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      kernelSyncOnlyRecalcSkips: 1,
      formulasParsed: 0,
      formulasBound: 0,
      lookupOwnerBuilds: 0,
    })
    expect(workbook.getStats().lastMetrics).toMatchObject({
      changedInputCount: 32,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
    })

    workbook.setCellContents(cell(sheetId, 0, 3), rowCount + 1_000)
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Number,
      value: rowCount,
    })
  })

  it('replaces literal sheet content in one undoable batch, including clears', () => {
    const workbook = WorkPaper.buildFromArray([
      [1, 2],
      [3, 4],
    ])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.setSheetContent(sheetId, [
      [10, 20],
      [null, 5],
    ])

    expect(changes).toHaveLength(4)
    expect(workbook.getCellSerialized(cell(sheetId, 0, 0))).toBe(10)
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe(20)
    expect(workbook.getCellSerialized(cell(sheetId, 1, 0))).toBeNull()
    expect(workbook.getCellSerialized(cell(sheetId, 1, 1))).toBe(5)

    const undoChanges = workbook.undo()

    expect(undoChanges).toHaveLength(4)
    expect(workbook.getSheetSerialized(sheetId)).toEqual([
      [1, 2],
      [3, 4],
    ])
  })

  it('replaces mixed sheet content in one undoable batch and binds formulas against loaded literals', () => {
    const workbook = WorkPaper.buildFromArray([[0]])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.setSheetContent(sheetId, [
      [1, 10, '=A1+B1'],
      [2, 20, '=A2+B2'],
    ])

    expect(changes).toHaveLength(6)
    expect(workbook.getCellFormula(cell(sheetId, 0, 2))).toBe('=A1+B1')
    expect(workbook.getCellValue(cell(sheetId, 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: 22,
    })

    const undoChanges = workbook.undo()

    expect(undoChanges).toHaveLength(6)
    expect(workbook.getSheetSerialized(sheetId)).toEqual([[0]])
  })

  it('keeps deferred literal batch updates correct across multiple sheets', () => {
    const workbook = WorkPaper.buildFromSheets({
      First: [[1], ['=A1*2']],
      Second: [[3]],
    })
    const firstId = workbook.getSheetId('First')!
    const secondId = workbook.getSheetId('Second')!

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(firstId, 0, 0), 10)
      workbook.setCellContents(cell(secondId, 0, 0), 7)
    })

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
      'First!A1',
      'First!A2',
      'Second!A1',
    ])
    expect(workbook.getCellValue(cell(firstId, 1, 0))).toEqual({
      tag: ValueTag.Number,
      value: 20,
    })

    const undoChanges = workbook.undo()

    expect(undoChanges.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
      'First!A1',
      'First!A2',
      'Second!A1',
    ])
    expect(workbook.getCellValue(cell(firstId, 1, 0))).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
  })

  it('suppresses readable value getters while evaluation is suspended and flushes on resume', () => {
    const workbook = WorkPaper.buildFromArray([[1]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const events: string[] = []

    workbook.on('evaluationSuspended', () => {
      events.push('suspend')
    })
    workbook.on('evaluationResumed', (changes) => {
      events.push(`resume:${changes.length}`)
    })
    workbook.on('valuesUpdated', (changes) => {
      events.push(`values:${changes.length}`)
    })

    workbook.suspendEvaluation()
    workbook.setCellContents(cell(sheetId, 0, 1), '=A1+1')

    expect(() => workbook.getCellValue(cell(sheetId, 0, 1))).toThrow(WorkPaperEvaluationSuspendedError)

    const changes = workbook.resumeEvaluation()

    expect(changes).toHaveLength(1)
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(events).toEqual(['suspend', 'resume:1', 'values:1'])
  })

  it('defers suspended cell mutations until resume and flushes them as one undoable engine batch', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const beforeBatchId = workbook.getStats().lastMetrics.batchId
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('suspended resume on a fresh workbook should use tracked engine changes')
    })

    workbook.suspendEvaluation()
    workbook.setCellContents(cell(sheetId, 0, 1), '=A1+A2')
    workbook.setCellContents(cell(sheetId, 0, 0), 10)
    workbook.setCellContents(cell(sheetId, 1, 0), 20)

    expect(workbook.getStats().lastMetrics.batchId).toBe(beforeBatchId)

    const changes = workbook.resumeEvaluation()

    expect(workbook.getStats().lastMetrics.batchId).toBe(beforeBatchId + 1)
    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
      'Sheet1!A1',
      'Sheet1!B1',
      'Sheet1!A2',
    ])
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })
    captureVisibilitySnapshot.mockRestore()

    const undoChanges = workbook.undo()

    expect(undoChanges.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : ''))).toEqual([
      'Sheet1!A1',
      'Sheet1!B1',
      'Sheet1!A2',
    ])
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBeNull()
  })

  it('supports custom scalar functions and clipboard translation for pasted formulas', () => {
    WorkPaper.registerFunctionPlugin({
      id: 'custom-math',
      implementedFunctions: {
        DOUBLE: { method: 'DOUBLE' },
      },
      functions: {
        DOUBLE: (value) => {
          if (value?.tag !== ValueTag.Number) {
            return { tag: ValueTag.Error, code: 3 }
          }
          return { tag: ValueTag.Number, value: value.value * 2 }
        },
      },
    })

    const workbook = WorkPaper.buildFromArray([[2]])
    const sheetId = workbook.getSheetId('Sheet1')!

    workbook.setCellContents(cell(sheetId, 0, 1), '=DOUBLE(A1)')

    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(workbook.calculateFormula('=DOUBLE(3)')).toMatchObject({
      tag: ValueTag.Number,
      value: 6,
    })

    const copied = workbook.copy({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 0, 1),
    })
    expect(copied[0]?.[1]).toMatchObject({ tag: ValueTag.Number, value: 4 })

    workbook.paste(cell(sheetId, 1, 0))

    expect(workbook.getCellSerialized(cell(sheetId, 1, 0))).toBe(2)
    expect(workbook.getCellFormula(cell(sheetId, 1, 1))).toBe('=DOUBLE(A2)')
    expect(workbook.getCellValue(cell(sheetId, 1, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    })
  })

  it('evaluates scratch formulas without mutating workbook sheets or undo history', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [[2, 3]],
    })
    const beforeSheetNames = workbook.getSheetNames()
    const beforeCanUndo = workbook.isThereSomethingToUndo()

    expect(workbook.calculateFormula('=SUM(Sheet1!A1:B1)')).toMatchObject({
      tag: ValueTag.Number,
      value: 5,
    })

    expect(workbook.getSheetNames()).toEqual(beforeSheetNames)
    expect(workbook.isThereSomethingToUndo()).toBe(beforeCanUndo)
  })

  it('rebuilds engine state when config changes affect available function plugins', () => {
    const plugin = {
      id: 'custom-math',
      implementedFunctions: {
        DOUBLE: { method: 'DOUBLE' },
      },
      functions: {
        DOUBLE: (value) => {
          if (value?.tag !== ValueTag.Number) {
            return { tag: ValueTag.Error, code: ErrorCode.Value }
          }
          return { tag: ValueTag.Number, value: value.value * 2 }
        },
      },
    } as const

    WorkPaper.registerFunctionPlugin(plugin)

    const workbook = WorkPaper.buildFromArray([[2]], { functionPlugins: [plugin] })
    const sheetId = workbook.getSheetId('Sheet1')!

    workbook.setCellContents(cell(sheetId, 0, 1), '=DOUBLE(A1)')
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    })

    workbook.updateConfig({
      functionPlugins: [{ id: 'missing-plugin', implementedFunctions: {} }],
    })

    expect(workbook.getCellFormula(cell(sheetId, 0, 1))).toBe('=DOUBLE(A1)')
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
  })

  it('preserves workbook semantics across rebuildAndRecalculate and non-semantic config rebuilds', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [[1, '=A1*Rate'], [3], [5]],
        Summary: [['=FILTER(Data!A1:A3,Data!A1:A3>2)']],
      },
      {
        useStats: false,
      },
    )
    const dataId = workbook.getSheetId('Data')!
    const summaryId = workbook.getSheetId('Summary')!

    workbook.addNamedExpression('Rate', '=2')

    const beforeDataSerialized = workbook.getSheetSerialized(dataId)
    const beforeSummaryValues = workbook.getRangeValues({
      start: cell(summaryId, 0, 0),
      end: cell(summaryId, 1, 0),
    })
    const beforeRateValue = workbook.getNamedExpressionValue('Rate')
    const rebuildChanges = workbook.rebuildAndRecalculate()

    expect(rebuildChanges).toEqual([])
    expect(workbook.getSheetSerialized(dataId)).toEqual(beforeDataSerialized)
    expect(
      workbook.getRangeValues({
        start: cell(summaryId, 0, 0),
        end: cell(summaryId, 1, 0),
      }),
    ).toEqual(beforeSummaryValues)
    expect(workbook.getNamedExpressionValue('Rate')).toEqual(beforeRateValue)

    workbook.updateConfig({
      useColumnIndex: true,
      useStats: true,
    })

    expect(workbook.getCellFormula(cell(dataId, 0, 1))).toBe('=A1*Rate')
    expect(workbook.getCellValue(cell(dataId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(
      workbook.getRangeValues({
        start: cell(summaryId, 0, 0),
        end: cell(summaryId, 1, 0),
      }),
    ).toEqual(beforeSummaryValues)
    expect(workbook.getNamedExpressionValue('Rate')).toEqual(beforeRateValue)
  })

  it('exposes calculation settings through the public runtime surface and keeps persisted config in sync', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [[1, '=A1+1']],
      },
      {
        calculationSettings: { iterate: true, iterateCount: 10, iterateDelta: '0.1', calcOnSave: true, calcCompleted: false },
      },
    )

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 10,
      iterateDelta: '0.1',
      calcOnSave: true,
      calcCompleted: false,
    })
    expect(workbook.getConfig().calculationSettings).toEqual({
      iterate: true,
      iterateCount: 10,
      iterateDelta: '0.1',
      calcOnSave: true,
      calcCompleted: false,
    })

    workbook.setCalculationSettings({ iterate: true, iterateCount: 25, iterateDelta: '0.001', calcOnSave: true, calcCompleted: true })

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
      calcOnSave: true,
      calcCompleted: true,
    })
    expect(workbook.getConfig().calculationSettings).toEqual({
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
      calcOnSave: true,
      calcCompleted: true,
    })
    expect(exportWorkPaperDocument(workbook).config?.calculationSettings).toEqual({
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
      calcOnSave: true,
      calcCompleted: true,
    })
  })

  it('reapplies config calculation settings after snapshot imports', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Iterative Snapshot',
        metadata: {
          calculationSettings: {
            mode: 'automatic',
            compatibilityMode: 'excel-modern',
            iterate: false,
          },
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Revolver',
          order: 0,
          cells: [
            { address: 'A1', value: 'Metric' },
            { address: 'B1', value: 'Value' },
            { address: 'A2', value: 'Opening debt' },
            { address: 'B2', value: 100000 },
            { address: 'A3', value: 'Interest rate' },
            { address: 'B3', value: 0.1 },
            { address: 'A4', value: 'Cash available for debt service' },
            { address: 'B4', value: 5000 },
            { address: 'A5', value: 'Interest expense' },
            { address: 'B5', formula: '=B6*B3' },
            { address: 'A6', value: 'Ending debt' },
            { address: 'B6', formula: '=B2+B5-B4' },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot, {
      maxColumns: 8,
      maxRows: 32,
      useColumnIndex: true,
      calculationSettings: { iterate: true, iterateCount: 100, iterateDelta: '0.0000000001' },
    })
    const sheetId = workbook.getSheetId('Revolver')!

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 100,
      iterateDelta: '0.0000000001',
    })
    expect(workbook.getCellValue(cell(sheetId, 4, 1))).toMatchObject({ tag: ValueTag.Number })
    expect(workbook.getCellValue(cell(sheetId, 4, 1)).value).toBeCloseTo(10555.555555555555, 10)
    expect(workbook.getCellValue(cell(sheetId, 5, 1))).toMatchObject({ tag: ValueTag.Number })
    expect(workbook.getCellValue(cell(sheetId, 5, 1)).value).toBeCloseTo(105555.55555555555, 10)
    workbook.dispose()
  })

  it('preserves rebuilt calculation settings across snapshot-reuse config updates', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Revolver: [
          ['Metric', 'Value'],
          ['Opening debt', 100000],
          ['Interest rate', 0.1],
          ['Cash available for debt service', 5000],
          ['Interest expense', '=B6*B3'],
          ['Ending debt', '=B2+B5-B4'],
        ],
      },
      {
        maxColumns: 8,
        maxRows: 32,
        useColumnIndex: true,
        calculationSettings: { iterate: false },
      },
    )
    const sheetId = workbook.getSheetId('Revolver')!

    expect(workbook.getCellValue(cell(sheetId, 4, 1))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })
    expect(workbook.getCellValue(cell(sheetId, 5, 1))).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })

    workbook.updateConfig({
      maxColumns: 8,
      maxRows: 32,
      useColumnIndex: true,
      calculationSettings: { iterate: true, iterateCount: 100, iterateDelta: '0.0000000001' },
    })

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 100,
      iterateDelta: '0.0000000001',
    })
    expect(workbook.getCellValue(cell(sheetId, 4, 1))).toMatchObject({ tag: ValueTag.Number })
    expect(workbook.getCellValue(cell(sheetId, 4, 1)).value).toBeCloseTo(10555.555555555555, 10)
    expect(workbook.getCellValue(cell(sheetId, 5, 1))).toMatchObject({ tag: ValueTag.Number })
    expect(workbook.getCellValue(cell(sheetId, 5, 1)).value).toBeCloseTo(105555.55555555555, 10)
  })

  it('preserves undo history across runtime-only config toggles', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [[1, '=MATCH(3,A1:A3,0)'], [2], [3]],
      },
      { useColumnIndex: false, useStats: false },
    )
    const sheetId = workbook.getSheetId('Bench')!

    workbook.setCellContents(cell(sheetId, 0, 0), 2)
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(workbook.isThereSomethingToUndo()).toBe(true)

    workbook.updateConfig({ useColumnIndex: true, useStats: true })

    expect(workbook.getConfig()).toMatchObject({ useColumnIndex: true, useStats: true })
    expect(workbook.isThereSomethingToUndo()).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    const undoChanges = workbook.undo()
    expect(undoChanges).not.toEqual([])
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
  })

  it('applies useColumnIndex to rebuilt engines when mixed with rebuild-only config updates', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [[1, '=MATCH(3,A1:A3,0)'], [2], [3]],
      },
      { useColumnIndex: false, language: 'enGB' },
    )

    expect(readEngineUseColumnIndexEnabled(workbook)).toBe(false)

    workbook.updateConfig({ useColumnIndex: true, language: 'rebuilt-language' })

    expect(workbook.getConfig()).toMatchObject({ useColumnIndex: true, language: 'rebuilt-language' })
    expect(readEngineUseColumnIndexEnabled(workbook)).toBe(true)
    expect(workbook.getCellValue(cell(workbook.getSheetId('Bench')!, 0, 1))).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
  })

  it('changes workbook-scoped named expressions without full visibility or named-value snapshots when no listeners need them', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [[1, '=Rate+1', '=Rate*2'], [2]],
      },
      {},
      [{ name: 'Rate', expression: '=2' }],
    )
    const sheetId = workbook.getSheetId('Bench')!
    const visibilitySnapshots = trackPrivateMethod(workbook, 'captureVisibilitySnapshot')
    const namedValueSnapshots = trackPrivateMethod(workbook, 'captureNamedExpressionValueSnapshot')

    try {
      const changes = workbook.changeNamedExpression('Rate', '=3')

      expect(changes.map((change) => (change.kind === 'cell' ? `cell:${change.a1}` : `name:${change.name}`))).toEqual([
        'cell:B1',
        'cell:C1',
        'name:Rate',
      ])
      expect(changes[0]).toMatchObject({
        kind: 'cell',
        a1: 'B1',
        newValue: { tag: ValueTag.Number, value: 4 },
      })
      expect(changes[1]).toMatchObject({
        kind: 'cell',
        a1: 'C1',
        newValue: { tag: ValueTag.Number, value: 6 },
      })
      expect(changes[2]).toMatchObject({
        kind: 'named-expression',
        name: 'Rate',
        newValue: { tag: ValueTag.Number, value: 3 },
      })
      expect(workbook.getNamedExpressionValue('Rate')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 4 })
      expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(visibilitySnapshots.count).toBe(0)
      expect(namedValueSnapshots.count).toBe(0)
    } finally {
      visibilitySnapshots.restore()
      namedValueSnapshots.restore()
    }
  })

  it('keeps simple scalar named expression changes on the snapshot-free path', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Bench: [[1, '=Rate+1', '=Rate*2']],
      },
      {},
      [{ name: 'Rate', expression: '=2' }],
    )
    const calculateFormula = trackPrivateMethod(workbook, 'calculateFormula')
    workbook.resetPerformanceCounters()

    try {
      const changes = workbook.changeNamedExpression('Rate', '=3')

      expect(changes).toContainEqual({
        kind: 'named-expression',
        name: 'Rate',
        scope: undefined,
        newValue: { tag: ValueTag.Number, value: 3 },
      })
      expect(workbook.getNamedExpressionValue('Rate')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(workbook.getSheetId('Bench')!, 0, 1))).toEqual({ tag: ValueTag.Number, value: 4 })
      expect(workbook.getCellValue(cell(workbook.getSheetId('Bench')!, 0, 2))).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        formulasBound: 0,
        wasmFullUploads: 0,
      })
      expect(calculateFormula.count).toBe(0)
    } finally {
      calculateFormula.restore()
    }
  })

  it('renames sheets without visibility snapshots when the rename preserves formula values and no listeners need events', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1], [2], [3]],
      Summary: [['=Data!A1+1', '=SUM(Data!A1:A3)']],
    })
    const dataSheetId = workbook.getSheetId('Data')!
    const summarySheetId = workbook.getSheetId('Summary')!
    const visibilitySnapshots = trackPrivateMethod(workbook, 'captureVisibilitySnapshot')

    try {
      const changes = workbook.renameSheet(dataSheetId, 'Source')

      expect(changes).toEqual([])
      expect(workbook.getSheetNames()).toEqual(['Source', 'Summary'])
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 0))).toBe('=Source!A1+1')
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 1))).toBe('=SUM(Source!A1:A3)')
      expect(workbook.getCellValue(cell(summarySheetId, 0, 0))).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(workbook.getCellValue(cell(summarySheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(visibilitySnapshots.count).toBe(0)

      workbook.undo()
      expect(workbook.getSheetNames()).toEqual(['Data', 'Summary'])
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 0))).toBe('=Data!A1+1')
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 1))).toBe('=SUM(Data!A1:A3)')

      workbook.redo()
      expect(workbook.getSheetNames()).toEqual(['Source', 'Summary'])
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 0))).toBe('=Source!A1+1')
      expect(workbook.getCellFormula(cell(summarySheetId, 0, 1))).toBe('=SUM(Source!A1:A3)')
    } finally {
      visibilitySnapshots.restore()
    }
  })

  it('returns changes in deterministic order for cells and named expressions', () => {
    const workbook = WorkPaper.buildFromArray([[]])
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 1), 20)
      workbook.setCellContents(cell(sheetId, 0, 0), 10)
      workbook.addNamedExpression('Zulu', '=1')
      workbook.addNamedExpression('Alpha', '=2')
    })

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.kind}:${change.a1}` : `${change.kind}:${change.name}`))).toEqual([
      'cell:A1',
      'cell:B1',
      'named-expression:Alpha',
      'named-expression:Zulu',
    ])
  })

  it('supports once listeners, address formatting, range dependency helpers, and tuple axis operations', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1, 2, '=A1+B1']],
    })
    const sheetId = workbook.getSheetId('Data')!
    let valuesUpdatedEvents = 0

    workbook.once('valuesUpdated', () => {
      valuesUpdatedEvents += 1
    })

    expect(workbook.simpleCellAddressToString(cell(sheetId, 0, 2))).toBe('C1')
    expect(workbook.simpleCellAddressToString(cell(sheetId, 0, 2), { includeSheetName: true })).toBe('Data!C1')

    expect(workbook.getCellDependents({ start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 1) })).toContainEqual({
      kind: 'cell',
      address: cell(sheetId, 0, 2),
    })

    workbook.setCellContents(cell(sheetId, 1, 0), 10)
    workbook.setCellContents(cell(sheetId, 1, 1), 20)

    expect(valuesUpdatedEvents).toBe(1)

    workbook.addRows(sheetId, [1, 1])
    expect(workbook.getSheetDimensions(sheetId).height).toBe(3)

    workbook.swapColumnIndexes(sheetId, [[0, 1]])
    expect(workbook.getCellSerialized(cell(sheetId, 0, 0))).toBe(2)
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe(1)
  })

  it('skips topo repair when appending independent direct aggregate formula rows', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const engine = Reflect.get(workbook, 'engine')
    if (typeof engine !== 'object' || engine === null || typeof Reflect.get(engine, 'resetPerformanceCounters') !== 'function') {
      throw new Error('Expected WorkPaper to expose an engine with performance counters in tests')
    }
    const applyCellMutationsAt = vi.spyOn(engineApplyCellMutationsTarget(workbook), 'applyCellMutationsAtWithOptions')

    try {
      Reflect.apply(Reflect.get(engine, 'resetPerformanceCounters'), engine, [])
      workbook.batch(() => {
        workbook.addRows(sheetId, 2, 2)
        workbook.setCellContents(cell(sheetId, 2, 0), [
          [5, 6, '=SUM(A3:B3)'],
          [7, 8, '=SUM(A4:B4)'],
        ])
      })

      const counters = Reflect.apply(Reflect.get(engine, 'getPerformanceCounters'), engine, [])
      expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
      expect(workbook.getCellValue(cell(sheetId, 3, 2))).toEqual({ tag: ValueTag.Number, value: 15 })
      expect(applyCellMutationsAt).toHaveBeenCalledTimes(1)
      expect(applyCellMutationsAt.mock.calls[0]?.[1]).toMatchObject({
        potentialNewCells: 6,
        reuseRefs: true,
        source: 'local',
      })
      expect(counters.topoRepairs).toBe(0)
      expect(counters.cycleFormulaScans).toBe(0)
      expect(counters.calcChainFullScans).toBe(0)
    } finally {
      applyCellMutationsAt.mockRestore()
    }
  })

  it('keeps topo repair when an appended formula depends on another formula', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, '=A1+1'],
        [3, 4],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    const engine = Reflect.get(workbook, 'engine')
    if (typeof engine !== 'object' || engine === null || typeof Reflect.get(engine, 'resetPerformanceCounters') !== 'function') {
      throw new Error('Expected WorkPaper to expose an engine with performance counters in tests')
    }

    Reflect.apply(Reflect.get(engine, 'resetPerformanceCounters'), engine, [])
    workbook.batch(() => {
      workbook.addRows(sheetId, 2, 1)
      workbook.setCellContents(cell(sheetId, 2, 0), [[5, '=B1+1']])
    })

    const counters = Reflect.apply(Reflect.get(engine, 'getPerformanceCounters'), engine, [])
    expect(workbook.getCellValue(cell(sheetId, 2, 1))).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(counters.topoRepairs).toBe(1)
  })

  it('uses HyperFormula-like optional returns for missing lookups and formula grids', () => {
    const workbook = WorkPaper.buildFromArray([[1, '=A1+1']])
    const sheetId = workbook.getSheetId('Sheet1')!

    expect(workbook.getSheetId('Missing')).toBeUndefined()
    expect(workbook.getSheetName(99)).toBeUndefined()
    expect(workbook.simpleCellAddressFromString('not-an-address')).toBeUndefined()
    expect(workbook.simpleCellRangeFromString('A1')).toBeUndefined()
    expect(workbook.getNamedExpression('Missing')).toBeUndefined()
    expect(workbook.getNamedExpressionFormula('Missing')).toBeUndefined()
    expect(workbook.getNamedExpressionValue('Missing')).toBeUndefined()
    expect(workbook.getRangeFormulas({ start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 1) })).toEqual([[undefined, '=A1+1']])
  })

  it('returns no value changes for structural row inserts when repeated direct aggregates preserve values', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, '=SUM(A1:A1)'],
        [2, '=SUM(A1:A2)'],
        [3, '=SUM(A1:A3)'],
        [4, '=SUM(A1:A4)'],
      ],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot')
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      const changes = workbook.addRows(sheetId, [1, 1])

      expect(changes).toEqual([])
      expect(captureVisibilitySnapshot).not.toHaveBeenCalled()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 5 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe('=SUM(A1:A1)')
      expect(workbook.getCellSerialized(cell(sheetId, 2, 1))).toBe('=SUM(A1:A3)')
      expect(workbook.getCellSerialized(cell(sheetId, 4, 1))).toBe('=SUM(A1:A5)')
      expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(workbook.getCellValue(cell(sheetId, 2, 1))).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(sheetId, 4, 1))).toEqual({ tag: ValueTag.Number, value: 10 })
    } finally {
      captureVisibilitySnapshot.mockRestore()
      forEachCellEntry.mockRestore()
    }
  })

  it('retargets owned direct aggregate row inserts without the generic reference scan', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: Array.from({ length: 16 }, (_value, row) => [row + 1, `=SUM(A1:A${row + 1})`]),
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const binding = readFormulaBindingTestSurface(workbook)
    const referenceScan = vi.spyOn(binding, 'collectFormulaCellsReferencingSheetNow')

    try {
      expect(workbook.addRows(sheetId, 8, 1)).toEqual([])
      expect(referenceScan).not.toHaveBeenCalled()

      expect(workbook.getCellSerialized(cell(sheetId, 16, 1))).toBe('=SUM(A1:A17)')
      expect(workbook.getCellValue(cell(sheetId, 16, 1))).toEqual({ tag: ValueTag.Number, value: 136 })

      workbook.setCellContents(cell(sheetId, 8, 0), 1000)
      expect(workbook.getCellValue(cell(sheetId, 9, 1))).toEqual({ tag: ValueTag.Number, value: 1045 })
    } finally {
      referenceScan.mockRestore()
    }
  })

  it('keeps simple formula column deletes quiet when dependencies survive', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, undefined, 2, '=A1+C1'],
        [3, undefined, 4, '=A2+C2'],
      ],
    })
    const sheetId = workbook.getSheetId('Sheet1')!

    const changes = workbook.removeColumns(sheetId, 1, 1)

    expect(changes).toEqual([])
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 2))).toBe('=A1+B1')
    expect(workbook.getCellSerialized(cell(sheetId, 1, 2))).toBe('=A2+B2')
    expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(workbook.getCellValue(cell(sheetId, 1, 2))).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('deletes repeated direct aggregate rows without visibility snapshots or dirty traversal', () => {
    const rowCount = 256
    const deleteRow = 127
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: Array.from({ length: rowCount }, (_, row) => [row + 1, `=SUM(A1:A${row + 1})`]),
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot')
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      workbook.resetPerformanceCounters()

      const changes = workbook.removeRows(sheetId, [deleteRow, 1])

      expect(changes).toEqual([])
      expect(captureVisibilitySnapshot).not.toHaveBeenCalled()
      forEachCellEntry.mockClear()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: rowCount - 1 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellValue(cell(sheetId, rowCount - 2, 1))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount * (rowCount + 1)) / 2 - (deleteRow + 1),
      })
      expect(workbook.getStats().lastMetrics).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        kernelSyncOnlyRecalcSkips: 1,
        regionQueryIndexBuilds: 0,
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
      forEachCellEntry.mockRestore()
    }
  })

  it('returns no value changes for structural column inserts when repeated simple families preserve values', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, 2, '=A1+B1', '=C1*2'],
        [2, 4, '=A2+B2', '=C2*2'],
        [3, 6, '=A3+B3', '=C3*2'],
      ],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const captureTracker = trackCaptureTrackedChangesWithoutVisibilityCache(workbook)
    const computeTrackedChanges = trackComputeCellChangesFromTrackedEvents(workbook)
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      const changes = workbook.addColumns(sheetId, [1, 1])

      expect(changes).toEqual([])
      expect(captureTracker.count).toBe(0)
      expect(computeTrackedChanges.count).toBe(0)
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 5, height: 3 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellSerialized(cell(sheetId, 0, 3))).toBe('=A1+C1')
      expect(workbook.getCellSerialized(cell(sheetId, 0, 4))).toBe('=D1*2')
      expect(workbook.getCellSerialized(cell(sheetId, 2, 3))).toBe('=A3+C3')
      expect(workbook.getCellSerialized(cell(sheetId, 2, 4))).toBe('=D3*2')
      expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(workbook.getCellValue(cell(sheetId, 2, 3))).toEqual({ tag: ValueTag.Number, value: 9 })
      expect(workbook.getCellValue(cell(sheetId, 2, 4))).toEqual({ tag: ValueTag.Number, value: 18 })
    } finally {
      forEachCellEntry.mockRestore()
      captureTracker.restore()
    }

    const undoChanges = workbook.undo()
    expect(undoChanges).toEqual([])
    expect(computeTrackedChanges.count).toBe(0)
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 4, height: 3 })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 2))).toBe('=A1+B1')
    expect(workbook.getCellSerialized(cell(sheetId, 0, 3))).toBe('=C1*2')
    computeTrackedChanges.restore()
  })

  it('defers simple formula families on first structural column insert without materializing the family index', () => {
    const rows = Array.from({ length: 48 }, (_, index) => {
      const rowNumber = index + 1
      return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=C${rowNumber}*2`]
    })
    const workbook = WorkPaper.buildFromSheets({ Sheet1: rows })
    const sheetId = workbook.getSheetId('Sheet1')!
    const binding = engineFormulaBindingTarget(workbook)
    const inspectFamilies = vi.spyOn(binding, 'forEachFormulaFamilyNow')
    const scanOwnedFormulas = vi.spyOn(binding, 'forEachFormulaCellOwnedBySheetNow')

    try {
      const changes = workbook.addColumns(sheetId, [1, 1])

      expect(changes).toEqual([])
      expect(inspectFamilies).not.toHaveBeenCalled()
      expect(scanOwnedFormulas).not.toHaveBeenCalled()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 5, height: 48 })
      expect(workbook.getCellSerialized(cell(sheetId, 47, 3))).toBe('=A48+C48')
      expect(workbook.getCellSerialized(cell(sheetId, 47, 4))).toBe('=D48*2')
      expect(workbook.getCellValue(cell(sheetId, 47, 4))).toEqual({ tag: ValueTag.Number, value: 288 })
    } finally {
      scanOwnedFormulas.mockRestore()
      inspectFamilies.mockRestore()
    }
  })

  it('composes repeated simple formula column inserts without materializing deferred families', () => {
    const rows = Array.from({ length: 48 }, (_, index) => {
      const rowNumber = index + 1
      return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=C${rowNumber}*2`]
    })
    const workbook = WorkPaper.buildFromSheets({ Sheet1: rows })
    const sheetId = workbook.getSheetId('Sheet1')!
    const binding = engineFormulaBindingTarget(workbook)
    const rebindFormulas = vi.spyOn(binding, 'rebindFormulaCellsNow')
    const inspectFamilies = vi.spyOn(binding, 'forEachFormulaFamilyNow')
    const scanOwnedFormulas = vi.spyOn(binding, 'forEachFormulaCellOwnedBySheetNow')

    try {
      expect(workbook.addColumns(sheetId, 1, 1)).toEqual([])
      expect(workbook.addColumns(sheetId, 1, 1)).toEqual([])

      expect(rebindFormulas).not.toHaveBeenCalled()
      expect(inspectFamilies).not.toHaveBeenCalled()
      expect(scanOwnedFormulas).not.toHaveBeenCalled()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 6, height: 48 })
      expect(workbook.getCellSerialized(cell(sheetId, 47, 4))).toBe('=A48+D48')
      expect(workbook.getCellSerialized(cell(sheetId, 47, 5))).toBe('=E48*2')
      expect(workbook.getCellValue(cell(sheetId, 47, 5))).toEqual({ tag: ValueTag.Number, value: 288 })
    } finally {
      scanOwnedFormulas.mockRestore()
      inspectFamilies.mockRestore()
      rebindFormulas.mockRestore()
    }
  })

  it('updates cached dimensions for safe middle column deletes without scanning the grid', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, 2, 3, 4],
        [5, 6, 7, 8],
      ],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      workbook.removeColumns(sheetId, [1, 1])

      forEachCellEntry.mockClear()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(sheetId, 1, 2))).toEqual({ tag: ValueTag.Number, value: 8 })
    } finally {
      forEachCellEntry.mockRestore()
    }
  })

  it('preserves cached dimensions for safe middle row moves without scanning the grid', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [[1], [2], [3], [4]],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      workbook.moveRows(sheetId, 1, 1, 0)

      forEachCellEntry.mockClear()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 1, height: 4 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(workbook.getCellValue(cell(sheetId, 1, 0))).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(workbook.getCellValue(cell(sheetId, 3, 0))).toEqual({ tag: ValueTag.Number, value: 4 })
    } finally {
      forEachCellEntry.mockRestore()
    }
  })

  it('preserves cached dimensions for safe middle column moves without scanning the grid', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [[1, 2, 3, 4]],
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const forEachCellEntry = vi.spyOn(sheetGridEntryTarget(workbook, sheetId), 'forEachCellEntry')

    try {
      workbook.moveColumns(sheetId, 1, 1, 0)

      forEachCellEntry.mockClear()
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 4, height: 1 })
      expect(forEachCellEntry).not.toHaveBeenCalled()
      expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({ tag: ValueTag.Number, value: 4 })
    } finally {
      forEachCellEntry.mockRestore()
    }
  })

  it('applies function translations to registered languages and exposes license validity', () => {
    WorkPaper.registerLanguage(TEST_LANGUAGE_CODE, { functions: {} })
    WorkPaper.registerFunctionPlugin(
      {
        id: 'custom-math',
        implementedFunctions: {
          DOUBLE: { method: 'DOUBLE' },
        },
      },
      {
        [TEST_LANGUAGE_CODE]: {
          DOUBLE: 'DUPLO',
        },
      },
    )

    expect(WorkPaper.getRegisteredFunctionNames(TEST_LANGUAGE_CODE)).toContain('DUPLO')
    expect(WorkPaper.buildEmpty().licenseKeyValidityState).toBe('valid')
    expect(WorkPaper.buildEmpty({ licenseKey: '' }).licenseKeyValidityState).toBe('missing')
  })
})
