import { describe, expect, it, vi } from 'vitest'
import * as formula from '@bilig/formula'
import { CellFlags, CellStore, readRuntimeImage, readRuntimeSnapshot, SpreadsheetEngine, WorkbookStore } from '@bilig/core'
import { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'
import { prepareInitialMixedSheetLoad } from '../initial-sheet-load.js'
import { WorkPaperSheetSizeLimitExceededError } from '../work-paper-errors.js'

describe('initial mixed sheet load', () => {
  it('builds mixed sheets without routing formulas through restore cell mutations', () => {
    const restoreMutationSpy = vi.spyOn(SpreadsheetEngine.prototype, 'applyCellMutationsAtWithOptions')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, '=A1*2'],
          [2, '=A2*3'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 2,
      })
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 6,
      })
      expect(restoreMutationSpy).not.toHaveBeenCalled()
    } finally {
      restoreMutationSpy.mockRestore()
    }
  })

  it('hydrates imported cached formula values when full calculation on load is disabled', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Imported',
        metadata: {
          calculationSettings: {
            mode: 'automatic',
            compatibilityMode: 'excel-modern',
            fullCalcOnLoad: false,
          },
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            { address: 'A1', value: 1 },
            { address: 'B1', formula: 'A1+1', value: 999 },
          ],
        },
      ],
    }

    const workbook = WorkPaper.buildFromSnapshot(snapshot)

    expect(workbook.getCellValue({ sheet: 1, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 999,
    })
  })

  it('normalizes repeated row-template formulas during mixed-sheet initialization', () => {
    const compileSpy = vi.spyOn(formula, 'compileFormulaAst')
    const parseSpy = vi.spyOn(formula, 'parseFormula')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 2, '=A1+B1', '=C1*2'],
          [2, 4, '=A2+B2', '=C2*2'],
          [3, 6, '=A3+B3', '=C3*2'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 3,
      })
      expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 18,
      })
      expect(workbook.getPerformanceCounters().formulasParsed).toBe(2)
      expect(workbook.getPerformanceCounters().directFormulaInitialEvaluations).toBe(6)
      expect(compileSpy).not.toHaveBeenCalled()
      expect(parseSpy).not.toHaveBeenCalled()
    } finally {
      compileSpy.mockRestore()
      parseSpy.mockRestore()
    }
  })

  it('normalizes row-template formulas with row-literal offsets during mixed-sheet initialization', () => {
    const compileSpy = vi.spyOn(formula, 'compileFormulaAst')
    const parseSpy = vi.spyOn(formula, 'parseFormula')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 2, '=A1+B1+1', '=C1*2+1'],
          [2, 4, '=A2+B2+2', '=C2*2+2'],
          [3, 6, '=A3+B3+3', '=C3*2+3'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 4,
      })
      expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 27,
      })
      expect(workbook.getPerformanceCounters().formulasParsed).toBe(2)
      expect(workbook.getPerformanceCounters().directFormulaInitialEvaluations).toBe(6)
      expect(compileSpy).not.toHaveBeenCalled()
      expect(parseSpy).not.toHaveBeenCalled()
    } finally {
      compileSpy.mockRestore()
      parseSpy.mockRestore()
    }
  })

  it('reuses inline initial direct scalar values when prefix aggregates require fallback evaluation', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, '=A1+1', '=SUM(B1:B1)'],
        [2, '=A2+1', '=SUM(B1:B2)'],
        [3, '=A3+1', '=SUM(B1:B3)'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 9,
    })
    expect(workbook.getPerformanceCounters().directFormulaInitialEvaluations).toBe(6)
  })

  it('re-evaluates inline direct scalars that depend on prefix aggregate formulas', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1+1', '=C1*2+1', '=SUM(A1:A1)+1', '=D1+E1+1'],
        [2, 4, '=A2+B2+2', '=C2*2+2', '=SUM(A1:A2)+2', '=D2+E2+2'],
        [3, 6, '=A3+B3+3', '=C3*2+3', '=SUM(A1:A3)+3', '=D3+E3+3'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 27,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 9,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 5 })).toEqual({
      tag: ValueTag.Number,
      value: 39,
    })
  })

  it('preserves large hydrated formula families for structural column inserts', () => {
    const rowCount = 3_000
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: rowCount }, (_value, index) => {
        const row = index + 1
        return [row, row * 2, `=A${row}+B${row}`, `=C${row}*2`]
      }),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const engine = Reflect.get(workbook, 'engine')
    const runtime = typeof engine === 'object' && engine !== null ? Reflect.get(engine, 'runtime') : undefined
    const binding = typeof runtime === 'object' && runtime !== null ? Reflect.get(runtime, 'binding') : undefined
    if (
      typeof binding !== 'object' ||
      binding === null ||
      typeof Reflect.get(binding, 'forEachFormulaCellOwnedBySheetNow') !== 'function'
    ) {
      throw new Error('Expected WorkPaper runtime binding service in test')
    }
    const ownedFormulaScan = vi.spyOn(binding, 'forEachFormulaCellOwnedBySheetNow')
    const familyScan = vi.spyOn(binding, 'forEachFormulaFamilyNow')

    try {
      workbook.resetPerformanceCounters()
      workbook.addColumns(sheetId, 1, 1)

      expect(workbook.getCellValue({ sheet: sheetId, row: rowCount - 1, col: 4 })).toEqual({
        tag: ValueTag.Number,
        value: rowCount * 6,
      })
      expect(ownedFormulaScan).not.toHaveBeenCalled()
      expect(familyScan.mock.calls.length).toBeLessThanOrEqual(1)
      expect(workbook.getPerformanceCounters()).toMatchObject({
        formulasBound: 0,
        structuralFormulaImpactCandidates: 0,
        structuralFormulaRebindInputs: 0,
      })
    } finally {
      ownedFormulaScan.mockRestore()
      familyScan.mockRestore()
      workbook.dispose()
    }
  })

  it('reserves mixed-sheet formula refs and attaches fresh cells without public per-cell attach calls', () => {
    const attachSpy = vi.spyOn(WorkbookStore.prototype, 'attachAllocatedCellWithLogicalAxisIds')
    const allocateReservedSpy = vi.spyOn(CellStore.prototype, 'allocateReserved')
    const initSpy = vi.spyOn(SpreadsheetEngine.prototype, 'initializeFormulaSourcesAtNow')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 10, '=A1+B1', '=C1*2'],
          [2, 20, '=A2+B2', '=C2*2'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!
      const refs = initSpy.mock.calls[0]?.[0] ?? []
      const potentialNewCells = initSpy.mock.calls[0]?.[1]

      expect(refs).toHaveLength(4)
      expect(Array.isArray(refs)).toBe(false)
      const collectedRefs = Array.from({ length: refs.length }, (_, index) => ({ ...refs.at(index) }))
      expect(collectedRefs.every((ref) => typeof ref.cellIndex === 'number')).toBe(true)
      expect(collectedRefs.map((ref) => ref.source)).toEqual(['A1+B1', 'C1*2', 'A2+B2', 'C2*2'])
      expect(potentialNewCells).toBe(0)
      expect(attachSpy).not.toHaveBeenCalled()
      expect(allocateReservedSpy).not.toHaveBeenCalled()
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 44,
      })
    } finally {
      attachSpy.mockRestore()
      allocateReservedSpy.mockRestore()
      initSpy.mockRestore()
    }
  })

  it('materializes dense mixed-sheet inspection rectangles without orphan cells', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'initial-mixed-dense-ragged-load' })
    engine.workbook.createSheet('Bench')
    const sheetId = engine.workbook.getSheet('Bench')!.id

    const prepared = prepareInitialMixedSheetLoad({
      engine,
      sheetId,
      content: [[1, '=A1*2'], [3]],
      rewriteFormula: (source) => source,
      inspection: {
        materializedCellCount: 4,
        maxColumnCount: 2,
        formulaCellCount: 1,
      },
    })

    const sheet = engine.workbook.getSheet('Bench')!
    const b2Index = engine.workbook.getCellIndex('Bench', 'B2')
    const row2Id = sheet.logicalAxisMap.getId('row', 1)!
    const colBId = sheet.logicalAxisMap.getId('column', 1)!

    expect(prepared.formulaRefs.length).toBe(1)
    expect(engine.workbook.cellStore.size).toBe(4)
    expect(engine.getCellValue('Bench', 'A2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Bench', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(b2Index).toBe(3)
    expect(sheet.grid.getPhysical(1, 1)).toBe(3)
    expect(sheet.logical.getCellVisiblePosition(3)).toEqual({ row: 1, col: 1 })
    expect(sheet.logical.listResidentCellIndices('row', [row2Id])).toEqual([2, 3])
    expect(sheet.logical.listResidentCellIndices('column', [colBId])).toEqual([1, 3])
    expect(engine.workbook.cellStore.flags[b2Index!] & CellFlags.AuthoredBlank).toBe(CellFlags.AuthoredBlank)
    expect(Array.from(sheet.columnVersions.slice(0, 2))).toEqual([1, 1])
  })

  it('bulk-clears dense mixed-sheet cell state before writing literals and formula refs', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'initial-mixed-dense-state-load' })
    engine.workbook.createSheet('Bench')
    const sheetId = engine.workbook.getSheet('Bench')!.id
    const cellStore = engine.workbook.cellStore
    cellStore.formulaIds.fill(77, 0, 4)
    cellStore.versions.fill(9, 0, 4)
    cellStore.topoRanks.fill(8, 0, 4)
    cellStore.cycleGroupIds.fill(7, 0, 4)
    cellStore.errors.fill(ErrorCode.Name, 0, 4)

    const prepared = prepareInitialMixedSheetLoad({
      engine,
      sheetId,
      content: [[1, 'label', '=A1+1', true]],
      rewriteFormula: (source) => source,
      inspection: {
        materializedCellCount: 4,
        maxColumnCount: 4,
        formulaCellCount: 1,
      },
    })

    expect(prepared.formulaRefs.length).toBe(1)
    expect(prepared.formulaRefs.at(0).cellIndex).toBe(2)
    expect(engine.getCellValue('Bench', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Bench', 'B1')).toMatchObject({ tag: ValueTag.String, value: 'label' })
    expect(engine.getCellValue('Bench', 'C1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Bench', 'D1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(Array.from(cellStore.formulaIds.slice(0, 4))).toEqual([0, 0, 0, 0])
    expect(Array.from(cellStore.errors.slice(0, 4))).toEqual([ErrorCode.None, ErrorCode.None, ErrorCode.None, ErrorCode.None])
    expect(Array.from(cellStore.versions.slice(0, 4))).toEqual([1, 1, 0, 1])
    expect(Array.from(cellStore.topoRanks.slice(0, 4))).toEqual([0, 0, 0, 0])
    expect(Array.from(cellStore.cycleGroupIds.slice(0, 4))).toEqual([-1, -1, -1, -1])

    engine.initializeFormulaSourcesAtNow(prepared.formulaRefs, prepared.potentialNewCells)

    expect(engine.getCellValue('Bench', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('merges compact initial formula refs across multiple mixed sheets', () => {
    const initSpy = vi.spyOn(SpreadsheetEngine.prototype, 'initializeFormulaSourcesAtNow')
    try {
      const workbook = WorkPaper.buildFromSheets({
        North: [
          [1, 10, '=A1+B1'],
          [2, 20, '=A2+B2'],
        ],
        South: [
          [3, 30, '=A1+B1'],
          [4, 40, '=A2+B2'],
        ],
      })
      const northId = workbook.getSheetId('North')!
      const southId = workbook.getSheetId('South')!
      const refs = initSpy.mock.calls[0]?.[0] ?? []

      expect(refs).toHaveLength(4)
      expect(Array.isArray(refs)).toBe(false)
      expect(Array.from({ length: refs.length }, (_, index) => refs.at(index)?.source)).toEqual(['A1+B1', 'A2+B2', 'A1+B1', 'A2+B2'])
      expect(Array.from({ length: refs.length }, (_, index) => typeof refs.at(index)?.cellIndex)).toEqual([
        'number',
        'number',
        'number',
        'number',
      ])
      expect(workbook.getCellValue({ sheet: northId, row: 1, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 22,
      })
      expect(workbook.getCellValue({ sheet: southId, row: 1, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 44,
      })
    } finally {
      initSpy.mockRestore()
    }
  })

  it('recognizes padded formulas without treating ordinary strings as formulas during mixed-sheet initialization', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[2, '  =A1*2  ', ' label ']],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellFormula({ sheet: sheetId, row: 0, col: 1 })).toBe('=A1*2')
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toMatchObject({
      tag: ValueTag.String,
      value: ' label ',
    })
  })

  it('hydrates fresh formula instance metadata without rescanning cell positions', () => {
    const positionSpy = vi.spyOn(WorkbookStore.prototype, 'getCellPosition')
    let workbook: WorkPaper | undefined
    try {
      workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 2, '=A1+B1', '=C1*2'],
          [2, 4, '=A2+B2', '=C2*2'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 12,
      })
      expect(positionSpy).not.toHaveBeenCalled()

      const runtimeImage = readRuntimeImage(readRuntimeSnapshot(workbook.getAllSheetsSerialized()))
      expect(
        runtimeImage?.formulaInstances.map(({ sheetName, row, col, source: formulaSource }) => ({
          sheetName,
          row,
          col,
          source: formulaSource,
        })),
      ).toEqual([
        { sheetName: 'Bench', row: 0, col: 2, source: 'A1+B1' },
        { sheetName: 'Bench', row: 0, col: 3, source: 'C1*2' },
        { sheetName: 'Bench', row: 1, col: 2, source: 'A2+B2' },
        { sheetName: 'Bench', row: 1, col: 3, source: 'C2*2' },
      ])
    } finally {
      positionSpy.mockRestore()
      workbook?.dispose()
    }
  })

  it('rebuilds from serialized sheets through the runtime-image fast path when available', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1', '=C1*2'],
        [2, 4, '=A2+B2', '=C2*2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    const runtimeImage = readRuntimeImage(readRuntimeSnapshot(serialized))
    expect(runtimeImage?.sheetCells?.[0]?.dimensions).toEqual({ width: 4, height: 2 })
    expect(runtimeImage?.sheetCells?.[0]?.cellCount).toBe(8)
    expect(runtimeImage?.formulaInstances).toHaveLength(4)
    expect(
      runtimeImage?.formulaInstances.map(({ sheetName, row, col, source: formulaSource }) => ({
        sheetName,
        row,
        col,
        source: formulaSource,
      })),
    ).toEqual([
      { sheetName: 'Bench', row: 0, col: 2, source: 'A1+B1' },
      { sheetName: 'Bench', row: 0, col: 3, source: 'C1*2' },
      { sheetName: 'Bench', row: 1, col: 2, source: 'A2+B2' },
      { sheetName: 'Bench', row: 1, col: 3, source: 'C2*2' },
    ])
    source.dispose()

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(rebuilt.getPerformanceCounters().snapshotOpsReplayed).toBe(0)
    expect(rebuilt.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(rebuilt.getPerformanceCounters().wasmFullUploads).toBe(0)
    expect(rebuilt.getPerformanceCounters().runtimeHydratedDirectScalarFastBindings).toBe(4)
    expect(rebuilt.getPerformanceCounters().formulaFamilyRuntimeRunFallbacks).toBe(0)

    rebuilt.setCellContents({ sheet: sheetId, row: 0, col: 0 }, 3)

    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
  })

  it('imports compatible runtime snapshots without reading serialized sheet matrix entries', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [2, 4, '=A2+B2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()
    serialized.Bench = serialized.Bench.map(
      (row) =>
        new Proxy(row, {
          get(target, property, receiver) {
            if (typeof property === 'string' && /^\d+$/.test(property)) {
              throw new Error('snapshot fast path should not read serialized cell values')
            }
            return Reflect.get(target, property, receiver)
          },
        }),
    )

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('imports compatible runtime snapshots without reading serialized sheet rows', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [2, 4, '=A2+B2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()
    const sheet = serialized.Bench
    expect(sheet).toBeDefined()
    serialized.Bench = new Proxy(sheet, {
      get(target, property, receiver) {
        if (property === 'length' || (typeof property === 'string' && /^\d+$/.test(property))) {
          throw new Error('snapshot fast path should not read serialized sheet rows')
        }
        return Reflect.get(target, property, receiver)
      },
    })

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('restores compatible runtime snapshots with batched column version notifications', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [2, 4, '=A2+B2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()
    const batchImpl: unknown = Reflect.get(WorkbookStore.prototype, 'withBatchedColumnVersionUpdates')
    const notifyImpl: unknown = Reflect.get(WorkbookStore.prototype, 'notifyColumnsWritten')
    if (typeof batchImpl !== 'function' || typeof notifyImpl !== 'function') {
      throw new Error('Expected WorkbookStore column-version methods in test')
    }
    let batchDepth = 0
    let notifyCount = 0
    let unbatchedNotifyCount = 0
    const batchSpy = vi
      .spyOn(WorkbookStore.prototype, 'withBatchedColumnVersionUpdates')
      .mockImplementation(function (this: WorkbookStore, execute) {
        batchDepth += 1
        try {
          return Reflect.apply(batchImpl, this, [execute])
        } finally {
          batchDepth -= 1
        }
      })
    const notifySpy = vi
      .spyOn(WorkbookStore.prototype, 'notifyColumnsWritten')
      .mockImplementation(function (this: WorkbookStore, sheetId, columns) {
        notifyCount += 1
        if (batchDepth === 0) {
          unbatchedNotifyCount += 1
        }
        return Reflect.apply(notifyImpl, this, [sheetId, columns])
      })
    try {
      const rebuilt = WorkPaper.buildFromSheets(serialized)
      const sheetId = rebuilt.getSheetId('Bench')!

      expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 6,
      })
      expect(notifyCount).toBeGreaterThan(0)
      expect(unbatchedNotifyCount).toBe(0)
      rebuilt.dispose()
    } finally {
      notifySpy.mockRestore()
      batchSpy.mockRestore()
    }
  })

  it('restores value-only snapshots with dense column metadata without per-range metadata writes', () => {
    const setColumnMetadataSpy = vi.spyOn(WorkbookStore.prototype, 'setColumnMetadata')
    try {
      const columnCount = 512
      const metadataPassCount = 3
      const snapshot: WorkbookSnapshot = {
        version: 1,
        workbook: { name: 'Imported value-only workbook' },
        sheets: [
          {
            id: 1,
            name: 'Imported',
            order: 0,
            metadata: {
              columns: Array.from({ length: columnCount }, (_, index) => ({
                id: `column-${index + 1}`,
                index,
              })),
              columnMetadata: Array.from({ length: columnCount * metadataPassCount }, (_, index) => {
                const start = index % columnCount
                return {
                  start,
                  count: 1,
                  size: 80 + Math.floor(index / columnCount),
                  hidden: start % 97 === 0,
                  customWidth: true,
                  styleIndex: start % 11,
                }
              }),
            },
            cells: Array.from({ length: 64 }, (_, index) => ({
              address: formula.formatAddress(Math.floor(index / 16), index % 16),
              value: index,
            })),
          },
        ],
      }

      const workbook = WorkPaper.buildFromSnapshot(snapshot, {
        evaluationTimeoutMs: 1_000,
        useColumnIndex: true,
      })
      const sheetId = workbook.getSheetId('Imported')!
      const restoredColumnMetadata = workbook.exportSnapshot().sheets[0]?.metadata?.columnMetadata ?? []

      expect(setColumnMetadataSpy).not.toHaveBeenCalled()
      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 0 })).toEqual({
        tag: ValueTag.Number,
        value: 0,
      })
      expect(restoredColumnMetadata).toContainEqual({
        start: 511,
        count: 1,
        size: 82,
        hidden: false,
        customWidth: true,
        styleIndex: 5,
      })
      workbook.dispose()
    } finally {
      setColumnMetadataSpy.mockRestore()
    }
  })

  it('rejects oversized sheets before importing compatible runtime snapshots', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2],
        [3, 4],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()

    expect(() => WorkPaper.buildFromSheets(serialized, { maxRows: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
    expect(() => WorkPaper.buildFromSheets(serialized, { maxColumns: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
  })
})
