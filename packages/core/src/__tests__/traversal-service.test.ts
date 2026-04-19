import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { makeCellEntity } from '../entity-ids.js'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineTraversalService } from '../engine/services/traversal-service.js'

function isEngineTraversalService(value: unknown): value is EngineTraversalService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'collectFormulaDependents') === 'function' &&
    typeof Reflect.get(value, 'forEachFormulaDependencyCell') === 'function' &&
    typeof Reflect.get(value, 'forEachSheetCell') === 'function'
  )
}

function getTraversalService(engine: SpreadsheetEngine): EngineTraversalService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const traversal = Reflect.get(runtime, 'traversal')
  if (!isEngineTraversalService(traversal)) {
    throw new TypeError('Expected engine traversal service')
  }
  return traversal
}

describe('EngineTraversalService', () => {
  it('collects formula dependents beyond the initial traversal scratch capacity', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'traversal-overflow' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    for (let row = 1; row <= 140; row += 1) {
      engine.setCellFormula('Sheet1', `B${row}`, 'A1+1')
    }
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B10)')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    expect(a1Index).toBeDefined()

    const dependents = Effect.runSync(getTraversalService(engine).collectFormulaDependents(makeCellEntity(a1Index!)))
    const dependentAddresses = [...dependents].map((cellIndex) => engine.workbook.getQualifiedAddress(cellIndex))

    expect(dependents.length).toBe(141)
    expect(new Set(dependentAddresses).size).toBe(141)
    expect(dependentAddresses).toContain('Sheet1!B1')
    expect(dependentAddresses).toContain('Sheet1!B140')
    expect(dependentAddresses).toContain('Sheet1!D1')
  })

  it('iterates formula dependencies and sheet cells through the extracted traversal boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'traversal-iteration' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'A1+2')
    engine.setCellValue('Sheet1', 'C2', 9)

    const traversal = getTraversalService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id

    expect(b1Index).toBeDefined()
    expect(sheetId).toBeDefined()

    const dependencyAddresses: string[] = []
    Effect.runSync(
      traversal.forEachFormulaDependencyCell(b1Index!, (dependencyCellIndex) => {
        dependencyAddresses.push(engine.workbook.getQualifiedAddress(dependencyCellIndex))
      }),
    )

    const sheetCells: string[] = []
    Effect.runSync(
      traversal.forEachSheetCell(sheetId!, (cellIndex, row, col) => {
        sheetCells.push(`${engine.workbook.getQualifiedAddress(cellIndex)}@${row},${col}`)
      }),
    )

    expect(dependencyAddresses).toEqual(['Sheet1!A1'])
    expect(sheetCells).toContain('Sheet1!A1@0,0')
    expect(sheetCells).toContain('Sheet1!B1@0,1')
    expect(sheetCells).toContain('Sheet1!C2@1,2')
  })

  it('exposes immediate traversal reads and wraps callback failures', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'traversal-direct-now' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'A1+2')

    const traversal = getTraversalService(engine)
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id

    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()
    expect(sheetId).toBeDefined()

    const dependents = traversal.getEntityDependentsNow(makeCellEntity(a1Index!))
    expect([...dependents].map((cellIndex) => engine.workbook.getQualifiedAddress(cellIndex))).toEqual(['Sheet1!B1'])

    const formulaDependents = traversal.collectFormulaDependentsNow(makeCellEntity(a1Index!))
    expect([...formulaDependents].map((cellIndex) => engine.workbook.getQualifiedAddress(cellIndex))).toEqual(['Sheet1!B1'])

    const directDependencies: string[] = []
    traversal.forEachFormulaDependencyCellNow(b1Index!, (dependencyCellIndex) => {
      directDependencies.push(engine.workbook.getQualifiedAddress(dependencyCellIndex))
    })
    expect(directDependencies).toEqual(['Sheet1!A1'])

    const directSheetCells: string[] = []
    traversal.forEachSheetCellNow(sheetId!, (cellIndex) => {
      directSheetCells.push(engine.workbook.getQualifiedAddress(cellIndex))
    })
    expect(directSheetCells).toEqual(['Sheet1!A1', 'Sheet1!B1'])

    expect(() =>
      Effect.runSync(
        traversal.forEachFormulaDependencyCell(b1Index!, () => {
          throw new Error('dependency boom')
        }),
      ),
    ).toThrow('dependency boom')
    expect(() =>
      Effect.runSync(
        traversal.forEachSheetCell(sheetId!, () => {
          throw new Error('sheet boom')
        }),
      ),
    ).toThrow('sheet boom')
  })

  it('collects exact and sorted lookup subscribers on the same column without entity collisions', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'traversal-lookup-collision',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellValue('Sheet1', 'D2', 25)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D2,A1:A3,1)')

    const a2Index = engine.workbook.getCellIndex('Sheet1', 'A2')
    expect(a2Index).toBeDefined()

    const dependents = Effect.runSync(getTraversalService(engine).collectFormulaDependents(makeCellEntity(a2Index!)))
    const dependentAddresses = [...dependents].map((cellIndex) => engine.workbook.getQualifiedAddress(cellIndex))

    expect(dependentAddresses).toContain('Sheet1!E1')
    expect(dependentAddresses).toContain('Sheet1!F1')
  })

  it('collects direct aggregate and criteria subscribers through symbolic region ownership', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'traversal-region-dependents' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'C1', 'SUMIF(A1:A4,">1",A1:A4)')

    const a2Index = engine.workbook.getCellIndex('Sheet1', 'A2')
    expect(a2Index).toBeDefined()

    const dependents = Effect.runSync(getTraversalService(engine).collectFormulaDependents(makeCellEntity(a2Index!)))
    const dependentAddresses = [...dependents].map((cellIndex) => engine.workbook.getQualifiedAddress(cellIndex))

    expect(dependentAddresses).toContain('Sheet1!B1')
    expect(dependentAddresses).toContain('Sheet1!C1')
  })
})
