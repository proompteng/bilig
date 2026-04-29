import { afterEach, describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'

import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from '../index.js'
import { WorkPaperEvaluationSuspendedError, WorkPaper } from '../index.js'
import { forceMaterializeTrackedIndexChanges, hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

const TEST_LANGUAGE_CODE = 'xHF'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

function expectOnlyCellChanges(changes: WorkPaperChange[]): asserts changes is WorkPaperCellChange[] {
  expect(changes.every((change) => change.kind === 'cell')).toBe(true)
}

function trackComputeCellChangesFromTrackedEvents(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, 'computeCellChangesFromTrackedEvents')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose computeCellChangesFromTrackedEvents in tests')
  }
  let count = 0
  Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, 'computeCellChangesFromTrackedEvents', original)
    },
  }
}

function rejectSingleTrackedCellReader(workbook: WorkPaper): { restore: () => void } {
  const original = Reflect.get(workbook, 'readSingleTrackedCellChange')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose readSingleTrackedCellChange in tests')
  }
  Reflect.set(workbook, 'readSingleTrackedCellChange', () => {
    throw new Error('Expected tiny sorted physical changes to avoid generic single-cell reading')
  })
  return {
    restore: () => {
      Reflect.set(workbook, 'readSingleTrackedCellChange', original)
    },
  }
}

function trackCaptureTrackedChangesWithoutVisibilityCache(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, 'captureTrackedChangesWithoutVisibilityCache')
  if (typeof original !== 'function') {
    throw new Error('Expected WorkPaper to expose captureTrackedChangesWithoutVisibilityCache in tests')
  }
  let count = 0
  Reflect.set(workbook, 'captureTrackedChangesWithoutVisibilityCache', (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, 'captureTrackedChangesWithoutVisibilityCache', original)
    },
  }
}

interface EngineApplyCellMutationsTarget {
  applyCellMutationsAtWithOptions: (...args: unknown[]) => unknown
}

interface SheetGridEntryTarget {
  forEachCellEntry: (fn: (cellIndex: number, row: number, col: number) => void) => void
}

interface SheetRecordTarget {
  grid: SheetGridEntryTarget
}

interface EngineWorkbookTarget {
  workbook: {
    getSheetById(sheetId: number): SheetRecordTarget | undefined
  }
}

function isEngineApplyCellMutationsTarget(value: unknown): value is EngineApplyCellMutationsTarget {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'applyCellMutationsAtWithOptions') === 'function'
}

function isEngineWorkbookTarget(value: unknown): value is EngineWorkbookTarget {
  const workbook = typeof value === 'object' && value !== null ? Reflect.get(value, 'workbook') : undefined
  return typeof workbook === 'object' && workbook !== null && typeof Reflect.get(workbook, 'getSheetById') === 'function'
}

function engineApplyCellMutationsTarget(workbook: WorkPaper): EngineApplyCellMutationsTarget {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineApplyCellMutationsTarget(engine)) {
    throw new Error('Expected WorkPaper to expose applyCellMutationsAtWithOptions in tests')
  }
  return engine
}

function sheetGridEntryTarget(workbook: WorkPaper, sheetId: number): SheetGridEntryTarget {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineWorkbookTarget(engine)) {
    throw new Error('Expected WorkPaper to expose workbook in tests')
  }
  const sheet = engine?.workbook?.getSheetById(sheetId)
  if (!sheet) {
    throw new Error('Expected WorkPaper to expose sheet grid in tests')
  }
  return sheet.grid
}

function readUndoStack(value: unknown): unknown[] | null {
  const engine = Reflect.get(value, 'engine')
  if (!engine || typeof engine !== 'object') {
    return null
  }
  const undoStack = Reflect.get(engine, 'undoStack')
  return Array.isArray(undoStack) ? undoStack : null
}

afterEach(() => {
  WorkPaper.unregisterAllFunctions()
  if (WorkPaper.getRegisteredLanguagesCodes().includes(TEST_LANGUAGE_CODE)) {
    WorkPaper.unregisterLanguage(TEST_LANGUAGE_CODE)
  }
})

describe('WorkPaper', () => {
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

  it('materializes no-listener tiny tracked changes eagerly without stale later writes', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[1, '=A1*2']],
    })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellContents(cell(sheetId, 0, 0), 9)

    expect(changes).toHaveLength(2)
    expectOnlyCellChanges(changes)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(false)

    workbook.setCellContents(cell(sheetId, 0, 0), 10)

    expect(changes[0]).toMatchObject({
      a1: 'A1',
      newValue: { tag: ValueTag.Number, value: 9 },
    })
    expect(changes[1]).toMatchObject({
      a1: 'B1',
      newValue: { tag: ValueTag.Number, value: 18 },
    })
    expect(forceMaterializeTrackedIndexChanges(changes)).toBe(false)
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
    } finally {
      genericReader.restore()
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

  it('keeps merged literal-only batch history on typed cell-mutation records', () => {
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
      'cell-mutations',
    )
    expect(Reflect.get(undoStack?.[0], 'inverse') ? Reflect.get(Reflect.get(undoStack?.[0], 'inverse'), 'kind') : undefined).toBe(
      'cell-mutations',
    )
  })

  it('passes known zero potential-new-cell count for existing-cell batch flushes', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!
    const applyCellMutationsAt = vi.spyOn(engineApplyCellMutationsTarget(workbook), 'applyCellMutationsAtWithOptions')

    try {
      workbook.batch(() => {
        workbook.setCellContents(cell(sheetId, 0, 0), 10)
        workbook.setCellContents(cell(sheetId, 1, 0), 20)
      })

      expect(applyCellMutationsAt).toHaveBeenCalledTimes(1)
      expect(applyCellMutationsAt.mock.calls[0]?.[1]).toMatchObject({
        captureUndo: true,
        potentialNewCells: 0,
        reuseRefs: true,
        source: 'local',
      })
    } finally {
      applyCellMutationsAt.mockRestore()
    }
  })

  it('passes known zero potential-new-cell count for existing-cell suspended flushes', () => {
    const workbook = WorkPaper.buildFromArray([[1], [2]])
    const sheetId = workbook.getSheetId('Sheet1')!

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
    } finally {
      applyCellMutationsAt.mockRestore()
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

    const changes = workbook.addRows(sheetId, [1, 1])

    expect(changes).toEqual([])
    expect(captureVisibilitySnapshot).not.toHaveBeenCalled()
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 5 })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe('=SUM(A1:A1)')
    expect(workbook.getCellSerialized(cell(sheetId, 2, 1))).toBe('=SUM(A1:A3)')
    expect(workbook.getCellSerialized(cell(sheetId, 4, 1))).toBe('=SUM(A1:A5)')
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(workbook.getCellValue(cell(sheetId, 2, 1))).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(workbook.getCellValue(cell(sheetId, 4, 1))).toEqual({ tag: ValueTag.Number, value: 10 })
    captureVisibilitySnapshot.mockRestore()
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

    const changes = workbook.addColumns(sheetId, [1, 1])

    expect(changes).toEqual([])
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 5, height: 3 })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 3))).toBe('=A1+C1')
    expect(workbook.getCellSerialized(cell(sheetId, 0, 4))).toBe('=D1*2')
    expect(workbook.getCellSerialized(cell(sheetId, 2, 3))).toBe('=A3+C3')
    expect(workbook.getCellSerialized(cell(sheetId, 2, 4))).toBe('=D3*2')
    expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(workbook.getCellValue(cell(sheetId, 2, 3))).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(workbook.getCellValue(cell(sheetId, 2, 4))).toEqual({ tag: ValueTag.Number, value: 18 })

    const undoChanges = workbook.undo()
    expect(undoChanges).toHaveLength(6)
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 4, height: 3 })
    expect(workbook.getCellSerialized(cell(sheetId, 0, 2))).toBe('=A1+B1')
    expect(workbook.getCellSerialized(cell(sheetId, 0, 3))).toBe('=C1*2')
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
