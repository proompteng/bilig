import { describe, expect, it, vi } from 'vitest'
import * as formula from '@bilig/formula'
import { readRuntimeImage, readRuntimeSnapshot, SpreadsheetEngine, WorkbookStore } from '@bilig/core'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { WorkPaper } from '../index.js'
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
      expect(familyScan).toHaveBeenCalledTimes(1)
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
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 44,
      })
    } finally {
      attachSpy.mockRestore()
      initSpy.mockRestore()
    }
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
