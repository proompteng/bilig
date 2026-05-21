import { afterEach, describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'

import type { WorkPaperCellAddress, WorkPaperCellChange, WorkPaperChange } from '../index.js'
import {
  WorkPaperEvaluationSuspendedError,
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '../index.js'
import { hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

const TEST_LANGUAGE_CODE = 'xHF'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function columnLabel(index: number): string {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

function trackPrivateMethod(workbook: WorkPaper, methodName: string): { readonly count: number; restore: () => void } {
  const original = Reflect.get(workbook, methodName)
  if (typeof original !== 'function') {
    throw new Error(`Expected WorkPaper to expose ${methodName} in tests`)
  }
  let count = 0
  Reflect.set(workbook, methodName, (...args: unknown[]) => {
    count += 1
    return Reflect.apply(original, workbook, args)
  })
  return {
    get count() {
      return count
    },
    restore: () => {
      Reflect.set(workbook, methodName, original)
    },
  }
}

interface TestSheetDimensionCache {
  updateAfterCellMutationRefs(...args: unknown[]): unknown
}

function hasSheetDimensionCacheUpdater(value: unknown): value is TestSheetDimensionCache {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'updateAfterCellMutationRefs') === 'function'
}

function trackSheetDimensionCacheUpdates(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const cache: unknown = Reflect.get(workbook, 'sheetDimensionCache')
  if (!hasSheetDimensionCacheUpdater(cache)) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  const spy = vi.spyOn(cache, 'updateAfterCellMutationRefs')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

function readEngineUseColumnIndexEnabled(workbook: WorkPaper): boolean {
  const engine = Reflect.get(workbook, 'engine')
  if (typeof engine !== 'object' || engine === null) {
    throw new Error('Expected WorkPaper to expose an engine object in tests')
  }
  return Reflect.get(engine, 'useColumnIndexEnabled') === true
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

interface EngineFormulaBindingTarget {
  runtime: {
    binding: {
      forEachFormulaFamilyNow: (fn: (...args: unknown[]) => void) => void
      isFormulaFamilyIndexReadyNow: () => boolean
    }
  }
}

function isEngineApplyCellMutationsTarget(value: unknown): value is EngineApplyCellMutationsTarget {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'applyCellMutationsAtWithOptions') === 'function'
}

function isEngineWorkbookTarget(value: unknown): value is EngineWorkbookTarget {
  const workbook = typeof value === 'object' && value !== null ? Reflect.get(value, 'workbook') : undefined
  return typeof workbook === 'object' && workbook !== null && typeof Reflect.get(workbook, 'getSheetById') === 'function'
}

function isEngineFormulaBindingTarget(value: unknown): value is EngineFormulaBindingTarget {
  const runtime = typeof value === 'object' && value !== null ? Reflect.get(value, 'runtime') : undefined
  const binding = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'binding') : undefined
  return (
    typeof binding === 'object' &&
    binding !== null &&
    typeof Reflect.get(binding, 'forEachFormulaFamilyNow') === 'function' &&
    typeof Reflect.get(binding, 'isFormulaFamilyIndexReadyNow') === 'function'
  )
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

function engineFormulaBindingTarget(workbook: WorkPaper): EngineFormulaBindingTarget['runtime']['binding'] {
  const engine = Reflect.get(workbook, 'engine')
  if (!isEngineFormulaBindingTarget(engine)) {
    throw new Error('Expected WorkPaper to expose formula binding service in tests')
  }
  return engine.runtime.binding
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
      'existing-numeric-cell-mutations',
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
        calculationSettings: { iterate: true, iterateCount: 10, iterateDelta: '0.1' },
      },
    )

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 10,
      iterateDelta: '0.1',
    })
    expect(workbook.getConfig().calculationSettings).toEqual({
      iterate: true,
      iterateCount: 10,
      iterateDelta: '0.1',
    })

    workbook.setCalculationSettings({ iterate: true, iterateCount: 25, iterateDelta: '0.001' })

    expect(workbook.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
    })
    expect(workbook.getConfig().calculationSettings).toEqual({
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
    })
    expect(exportWorkPaperDocument(workbook).config?.calculationSettings).toEqual({
      iterate: true,
      iterateCount: 25,
      iterateDelta: '0.001',
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
    expect(undoChanges).toHaveLength(6)
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
