import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { EngineFormulaBindingError } from '../engine/errors.js'
import type { EngineFormulaBindingService } from '../engine/services/formula-binding-service.js'
import type { FormulaFamilyStore } from '../formula/formula-family-store.js'

function isEngineFormulaBindingService(value: unknown): value is EngineFormulaBindingService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'bindFormula') === 'function' &&
    typeof Reflect.get(value, 'clearFormula') === 'function' &&
    typeof Reflect.get(value, 'rewriteCellFormulasForSheetRename') === 'function'
  )
}

function isFormulaFamilyStore(value: unknown): value is FormulaFamilyStore {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'getStats') === 'function' &&
    typeof Reflect.get(value, 'listFamilies') === 'function'
  )
}

function getBindingService(engine: SpreadsheetEngine): EngineFormulaBindingService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const binding = Reflect.get(runtime, 'binding')
  if (!isEngineFormulaBindingService(binding)) {
    throw new TypeError('Expected engine formula binding service')
  }
  return binding
}

function getFormulaFamilyStore(engine: SpreadsheetEngine): FormulaFamilyStore {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const formulaFamilies = Reflect.get(runtime, 'formulaFamilies')
  if (!isFormulaFamilyStore(formulaFamilies)) {
    throw new TypeError('Expected formula family store')
  }
  return formulaFamilies
}

function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null || typeof Reflect.get(formulas, 'get') !== 'function') {
    throw new TypeError('Expected internal formulas store')
  }
  return Reflect.get(formulas, 'get').call(formulas, cellIndex)
}

function readRuntimeTemplateId(engine: SpreadsheetEngine, cellIndex: number): number | undefined {
  const formula = readRuntimeFormula(engine, cellIndex)
  const templateId = typeof formula === 'object' && formula !== null ? Reflect.get(formula, 'templateId') : undefined
  return typeof templateId === 'number' ? templateId : undefined
}

function isRuntimeFormulaWithDirectCriteria(value: unknown): value is {
  directCriteria: {
    aggregateKind: string
    aggregateRange:
      | {
          sheetName: string
          rowStart: number
          rowEnd: number
          col: number
          length: number
        }
      | undefined
    criteriaPairs: Array<{
      range: { sheetName: string; rowStart: number; rowEnd: number; col: number; length: number }
      criterion: { kind: 'literal'; value: unknown } | { kind: 'cell'; cellIndex: number }
    }>
  }
} {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  const directCriteria = Reflect.get(value, 'directCriteria')
  return typeof directCriteria === 'object' && directCriteria !== null
}

describe('EngineFormulaBindingService', () => {
  it('indexes formula owners and qualified sheet references for structural candidate collection', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-sheet-indexes' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet2', 'C1', 'Sheet1!A1+1')

    const binding = getBindingService(engine)
    const sheet1Owned = binding.collectFormulaCellsOwnedBySheetNow('Sheet1')
    const sheet2Owned = binding.collectFormulaCellsOwnedBySheetNow('Sheet2')
    const sheet1Referenced = binding.collectFormulaCellsReferencingSheetNow('Sheet1')

    expect(sheet1Owned).toEqual([engine.workbook.getCellIndex('Sheet1', 'B1')!])
    expect(sheet2Owned).toEqual([engine.workbook.getCellIndex('Sheet2', 'C1')!])
    expect(sheet1Referenced).toEqual([engine.workbook.getCellIndex('Sheet2', 'C1')!])
  })

  it('clears reverse dependency edges when a formula is removed through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-clear' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(formulaCellIndex).toBeDefined()
    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toContain('Sheet1!B1')

    Effect.runSync(getBindingService(engine).clearFormula(formulaCellIndex!))

    expect(engine.getCell('Sheet1', 'B1').formula).toBeUndefined()
    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toEqual([])
  })

  it('rewrites quoted sheet references on rename through the binding service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-rename' })
    await engine.ready()
    engine.createSheet("Q1's Data")
    engine.createSheet('Summary')
    engine.setCellValue("Q1's Data", 'A1', 7)
    engine.setCellFormula('Summary', 'A1', "'Q1''s Data'!A1*2")

    const renamed = engine.workbook.renameSheet("Q1's Data", "Q2's Data")
    expect(renamed).toBeTruthy()

    Effect.runSync(getBindingService(engine).rewriteCellFormulasForSheetRename("Q1's Data", "Q2's Data", 0))

    expect(engine.getCell('Summary', 'A1').formula).toBe("'Q2''s Data'!A1*2")
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 14 })
  })

  it('binds repeated row-translated formulas through the service without changing results', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-row-template' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'B2', 20)

    engine.setCellFormula('Sheet1', 'E1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'F1', 'E1*2')
    engine.setCellFormula('Sheet1', 'E2', 'A2+B2')
    engine.setCellFormula('Sheet1', 'F2', 'E2*2')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 44 })
    expect(readRuntimeTemplateId(engine, engine.workbook.getCellIndex('Sheet1', 'E1')!)).toBeDefined()
    expect(readRuntimeTemplateId(engine, engine.workbook.getCellIndex('Sheet1', 'E2')!)).toBe(
      readRuntimeTemplateId(engine, engine.workbook.getCellIndex('Sheet1', 'E1')!),
    )
    expect(readRuntimeTemplateId(engine, engine.workbook.getCellIndex('Sheet1', 'F2')!)).toBe(
      readRuntimeTemplateId(engine, engine.workbook.getCellIndex('Sheet1', 'F1')!),
    )
  })

  it('registers repeated template formulas into compressed family runs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-formula-families' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 100; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    expect(getBindingService(engine).getFormulaFamilyStatsNow()).toEqual({
      familyCount: 4,
      runCount: 4,
      memberCount: 200,
    })
    expect(
      getFormulaFamilyStore(engine)
        .listFamilies()
        .flatMap((family) => family.runs.map((run) => run.cellIndices.length)),
    ).toEqual([1, 99, 1, 99])
  })

  it('runs tracked rebinding wrappers through the service surface', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-wrapper-rebinds' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const binding = getBindingService(engine)

    expect(Effect.runSync(binding.rebindDefinedNameDependents([], 3))).toBe(3)
    expect(Effect.runSync(binding.rebindTableDependents([], 5))).toBe(5)
    expect(Effect.runSync(binding.rebindFormulasForSheet('Sheet1', 7))).toBe(7)
    expect(Effect.runSync(binding.rebindFormulasForSheet('Sheet1', 11, []))).toBe(11)
  })

  it('preserves dependency wiring across formula rewrites with the same dependencies', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-same-deps-rewrite' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'C1*2')

    engine.resetPerformanceCounters()
    engine.setCellFormula('Sheet1', 'C1', 'A1*B1')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toContain('Sheet1!C1')
    expect(engine.getDependencies('Sheet1', 'C1').directDependents).toContain('Sheet1!D1')
    expect(engine.getPerformanceCounters()).toMatchObject({
      topoRepairs: 0,
      topoRepairAffectedFormulas: 0,
    })
  })

  it('binds direct criteria descriptors for supported conditional aggregate families', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-direct-criteria' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'B3', 30)
    engine.setCellValue('Sheet1', 'B4', 40)
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'F1', 'COUNTIF(A1:A4,">1")')
    engine.setCellFormula('Sheet1', 'F2', 'SUMIF(A1:A4,D1,B1:B4)')
    engine.setCellFormula('Sheet1', 'F3', 'AVERAGEIFS(B1:B4,A1:A4,D1)')

    const countIndex = engine.workbook.getCellIndex('Sheet1', 'F1')
    const sumIndex = engine.workbook.getCellIndex('Sheet1', 'F2')
    const averageIndex = engine.workbook.getCellIndex('Sheet1', 'F3')
    if (countIndex === undefined || sumIndex === undefined || averageIndex === undefined) {
      throw new Error('expected criteria formulas to be materialized')
    }

    const countFormula = readRuntimeFormula(engine, countIndex)
    if (!isRuntimeFormulaWithDirectCriteria(countFormula)) {
      throw new Error('expected COUNTIF runtime formula to expose direct criteria metadata')
    }
    expect(countFormula.directCriteria.aggregateKind).toBe('count')
    expect(countFormula.directCriteria.aggregateRange).toBeUndefined()
    expect(countFormula.directCriteria.criteriaPairs[0]?.criterion).toMatchObject({
      kind: 'literal',
    })

    const sumFormula = readRuntimeFormula(engine, sumIndex)
    if (!isRuntimeFormulaWithDirectCriteria(sumFormula)) {
      throw new Error('expected SUMIF runtime formula to expose direct criteria metadata')
    }
    expect(sumFormula.directCriteria.aggregateKind).toBe('sum')
    expect(sumFormula.directCriteria.aggregateRange).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 1,
      length: 4,
    })
    expect(sumFormula.directCriteria.criteriaPairs[0]?.criterion).toMatchObject({
      kind: 'cell',
    })

    const averageFormula = readRuntimeFormula(engine, averageIndex)
    if (!isRuntimeFormulaWithDirectCriteria(averageFormula)) {
      throw new Error('expected AVERAGEIFS runtime formula to expose direct criteria metadata')
    }
    expect(averageFormula.directCriteria.aggregateKind).toBe('average')
    expect(averageFormula.directCriteria.aggregateRange).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 1,
      length: 4,
    })
  })

  it('binds direct criteria descriptors for COUNTIFS SUMIFS MINIFS and MAXIFS', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-direct-criteria-families' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellValue('Sheet1', 'B1', 'x')
    engine.setCellValue('Sheet1', 'B2', 'x')
    engine.setCellValue('Sheet1', 'B3', 'y')
    engine.setCellValue('Sheet1', 'B4', 'x')
    engine.setCellValue('Sheet1', 'C1', 10)
    engine.setCellValue('Sheet1', 'C2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)
    engine.setCellValue('Sheet1', 'C4', 40)
    engine.setCellFormula('Sheet1', 'G1', 'COUNTIFS(A1:A4,">1",B1:B4,"x")')
    engine.setCellFormula('Sheet1', 'G2', 'SUMIFS(C1:C4,A1:A4,">1",B1:B4,"x")')
    engine.setCellFormula('Sheet1', 'G3', 'MINIFS(C1:C4,A1:A4,">1",B1:B4,"x")')
    engine.setCellFormula('Sheet1', 'G4', 'MAXIFS(C1:C4,A1:A4,">1",B1:B4,"x")')

    for (const address of ['G1', 'G2', 'G3', 'G4'] as const) {
      const cellIndex = engine.workbook.getCellIndex('Sheet1', address)
      if (cellIndex === undefined) {
        throw new Error(`expected ${address} to be materialized`)
      }
      const runtimeFormula = readRuntimeFormula(engine, cellIndex)
      if (!isRuntimeFormulaWithDirectCriteria(runtimeFormula)) {
        throw new Error(`expected ${address} to expose direct criteria metadata`)
      }
      expect(runtimeFormula.directCriteria.criteriaPairs).toHaveLength(2)
      expect(runtimeFormula.directCriteria.criteriaPairs[0]?.criterion).toMatchObject({
        kind: 'literal',
      })
    }

    const countFormula = readRuntimeFormula(engine, engine.workbook.getCellIndex('Sheet1', 'G1')!)
    const sumFormula = readRuntimeFormula(engine, engine.workbook.getCellIndex('Sheet1', 'G2')!)
    const minFormula = readRuntimeFormula(engine, engine.workbook.getCellIndex('Sheet1', 'G3')!)
    const maxFormula = readRuntimeFormula(engine, engine.workbook.getCellIndex('Sheet1', 'G4')!)
    if (
      !isRuntimeFormulaWithDirectCriteria(countFormula) ||
      !isRuntimeFormulaWithDirectCriteria(sumFormula) ||
      !isRuntimeFormulaWithDirectCriteria(minFormula) ||
      !isRuntimeFormulaWithDirectCriteria(maxFormula)
    ) {
      throw new Error('expected all supported criteria families to expose direct metadata')
    }
    expect(countFormula.directCriteria.aggregateKind).toBe('count')
    expect(sumFormula.directCriteria.aggregateKind).toBe('sum')
    expect(minFormula.directCriteria.aggregateKind).toBe('min')
    expect(maxFormula.directCriteria.aggregateKind).toBe('max')
  })

  it('does not bind direct criteria descriptors for unsupported criteria shapes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-direct-criteria-unsupported' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'B3', 30)
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'H2', 'SUMIF(A1:A4,D1,B1:B3)')
    engine.setCellFormula('Sheet1', 'H3', 'COUNTIFS(A1:A4,">1",B1:B3,"x")')

    for (const address of ['H2', 'H3'] as const) {
      const cellIndex = engine.workbook.getCellIndex('Sheet1', address)
      if (cellIndex === undefined) {
        throw new Error(`expected ${address} to be materialized`)
      }
      const runtimeFormula = readRuntimeFormula(engine, cellIndex)
      expect(isRuntimeFormulaWithDirectCriteria(runtimeFormula)).toBe(false)
    }
  })

  it('skips unmapped and missing-owner formulas during full rebuild', async () => {
    const missingMappingEngine = new SpreadsheetEngine({ workbookName: 'binding-rebuild-missing-mapping' })
    await missingMappingEngine.ready()
    missingMappingEngine.createSheet('Sheet1')
    missingMappingEngine.setCellValue('Sheet1', 'A1', 7)
    missingMappingEngine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const missingMappingIndex = missingMappingEngine.workbook.getCellIndex('Sheet1', 'B1')
    if (missingMappingIndex === undefined) {
      throw new Error('expected formula index for missing mapping case')
    }
    missingMappingEngine.workbook.cellStore.sheetIds[missingMappingIndex] = undefined

    const missingMappingResult = Effect.runSync(getBindingService(missingMappingEngine).rebuildAllFormulaBindings())
    expect(missingMappingResult).toEqual([])
    expect(readRuntimeFormula(missingMappingEngine, missingMappingIndex)).toBeUndefined()

    const missingOwnerEngine = new SpreadsheetEngine({ workbookName: 'binding-rebuild-missing-owner' })
    await missingOwnerEngine.ready()
    missingOwnerEngine.createSheet('Sheet1')
    missingOwnerEngine.setCellValue('Sheet1', 'A1', 7)
    missingOwnerEngine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const missingOwnerIndex = missingOwnerEngine.workbook.getCellIndex('Sheet1', 'B1')
    if (missingOwnerIndex === undefined) {
      throw new Error('expected formula index for missing owner case')
    }
    missingOwnerEngine.workbook.sheetsByName.delete('Sheet1')

    const missingOwnerResult = Effect.runSync(getBindingService(missingOwnerEngine).rebuildAllFormulaBindings())
    expect(missingOwnerResult).toEqual([])
    expect(readRuntimeFormula(missingOwnerEngine, missingOwnerIndex)).toBeUndefined()
  })

  it('invalidates formulas that fail to rebind during a full rebuild while preserving the active index list', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-rebuild-invalid-source' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    if (formulaIndex === undefined) {
      throw new Error('expected formula index')
    }
    const runtimeFormula = readRuntimeFormula(engine, formulaIndex)
    if (typeof runtimeFormula !== 'object' || runtimeFormula === null) {
      throw new Error('expected runtime formula to mutate source')
    }
    Reflect.set(runtimeFormula, 'source', 'SUM(')

    const rebound = Effect.runSync(getBindingService(engine).rebuildAllFormulaBindings())
    expect(rebound).toEqual([formulaIndex])
    expect(readRuntimeFormula(engine, formulaIndex)).toBeUndefined()
  })

  it('wraps rebuild and rebind failures with EngineFormulaBindingError', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-error-wrappers' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Other')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1+Other!A1')
    engine.setCellValue('Other', 'A1', 1)

    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    if (formulaIndex === undefined) {
      throw new Error('expected formula index')
    }

    const getSheetNameByIdSpy = vi.spyOn(engine.workbook, 'getSheetNameById').mockImplementation(() => {
      throw new Error('binding explode')
    })

    const binding = getBindingService(engine)
    for (const effect of [binding.rebuildAllFormulaBindings(), binding.rebindFormulaCells([formulaIndex], 0)]) {
      const result = Effect.runSync(Effect.either(effect))
      expect(result._tag).toBe('Left')
      expect(result.left).toBeInstanceOf(EngineFormulaBindingError)
      expect(result.left.message).toContain('binding explode')
    }

    getSheetNameByIdSpy.mockRestore()
  })

  it('wraps tracked dependent rebinding failures with EngineFormulaBindingError', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'binding-tracked-wrapper-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.createSheet('Other')
    engine.setCellFormula('Other', 'B1', 'Sheet1!A1*2')
    const formulaIndex = engine.workbook.getCellIndex('Other', 'B1')
    if (formulaIndex === undefined) {
      throw new Error('expected tracked formula indices')
    }
    const reverseTableEdges = Reflect.get(engine, 'reverseTableEdges')
    if (!(reverseTableEdges instanceof Map)) {
      throw new Error('expected table reverse-edge registry')
    }
    reverseTableEdges.set('Sales', new Set([formulaIndex]))

    const getSheetNameByIdSpy = vi.spyOn(engine.workbook, 'getSheetNameById').mockImplementation(() => {
      throw new Error('tracked explode')
    })

    const binding = getBindingService(engine)
    const definedNames = Effect.runSync(Effect.either(binding.rebindDefinedNameDependents([''], 0)))
    expect(definedNames._tag).toBe('Left')
    expect(definedNames.left).toBeInstanceOf(EngineFormulaBindingError)
    expect(definedNames.left.message).toContain('Defined names must be non-empty')

    const tableDependents = Effect.runSync(Effect.either(binding.rebindTableDependents(['Sales'], 0)))
    expect(tableDependents._tag).toBe('Left')
    expect(tableDependents.left).toBeInstanceOf(EngineFormulaBindingError)
    expect(tableDependents.left.message).toContain('tracked explode')

    const sheetRebind = Effect.runSync(Effect.either(binding.rebindFormulasForSheet('Sheet1', 0, [formulaIndex])))
    expect(sheetRebind._tag).toBe('Left')
    expect(sheetRebind.left).toBeInstanceOf(EngineFormulaBindingError)
    expect(sheetRebind.left.message).toContain('tracked explode')

    getSheetNameByIdSpy.mockRestore()
  })
})
