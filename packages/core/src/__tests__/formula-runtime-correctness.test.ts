import { describe, expect, it, vi } from 'vitest'
import { FormulaMode, ValueTag, type CellValue, type WorkbookDefinedNameValueSnapshot, type WorkbookTableSnapshot } from '@bilig/protocol'
import {
  canonicalFormulaFixtures,
  type ExcelExpectedValue,
  type ExcelFixtureCase,
  type ExcelFixtureTable,
} from '../../../excel-fixtures/src/index.js'
import { getCompatibilityEntry } from '../../../formula/src/compatibility.js'
import { SpreadsheetEngine } from '../index.js'

// These formulas still compile onto the wasm-capable path, but the engine intentionally
// reroutes them to specialized JS lookup handlers at bind time.
const runtimeJsOnlyFixtureIds = new Set(['lookup-reference:match-exact', 'lookup-reference:xmatch-basic'])

const engineRuntimeSkipReasons = new Map<string, string>([
  ['aggregation:avg-range', 'AVERAGE range aggregation still binds through the JS aggregate path in the engine.'],
  ['lookup-reference:offset-basic', 'OFFSET is contextual and verified through dedicated engine tests.'],
  ['names:defined-name-range', 'Defined-name range formulas require workbook metadata and are routed through JS.'],
  ['names:defined-name-scalar', 'Defined-name scalar formulas require workbook metadata and are routed through JS.'],
  ['structured-reference:table-column-ref', 'Structured references require workbook table metadata and are routed through JS.'],
  ['tables:table-total-row-sum', 'Table formulas require workbook table metadata and are routed through JS.'],
])
const capturedVolatileFixtureIds = new Set(['date-time:today-volatile', 'date-time:now-volatile', 'volatile:rand-basic'])
const capturedVolatileOracleTime = new Date('2026-03-19T15:45:30.000Z')

const engineRunnableProductionFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id)
  return (
    entry?.wasmStatus === 'production' &&
    !runtimeJsOnlyFixtureIds.has(fixture.id) &&
    !engineRuntimeSkipReasons.has(fixture.id) &&
    fixture.multipleOperations === undefined
  )
})

const groupedArrayProductionFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id)
  return (
    entry?.wasmStatus === 'production' && (fixture.id === 'dynamic-array:groupby-basic' || fixture.id === 'dynamic-array:pivotby-basic')
  )
})

const criteriaAggregateProductionFixtureIds = new Set([
  'statistical:averageif-basic',
  'statistical:averageifs-basic',
  'statistical:countif-basic',
  'statistical:countifs-basic',
  'statistical:sumif-basic',
  'statistical:sumifs-basic',
])

const criteriaAggregateProductionFixtures = canonicalFormulaFixtures.filter((fixture) => {
  const entry = getCompatibilityEntry(fixture.id)
  return entry?.wasmStatus === 'production' && criteriaAggregateProductionFixtureIds.has(fixture.id)
})

describe('formula runtime correctness', () => {
  it('keeps engine-runnable canonical production fixtures in oracle parity on the wasm path', async () => {
    expect(engineRunnableProductionFixtures.length).toBeGreaterThan(0)

    await Promise.all(
      engineRunnableProductionFixtures.map(async (fixture) => {
        try {
          await expectFixtureParity(fixture)
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          throw new Error(`Fixture ${fixture.id} failed: ${message}`, { cause: error })
        }
      }),
    )
  })

  it('documents every production fixture excluded from engine differential parity', () => {
    const skipped = canonicalFormulaFixtures.filter((fixture) => {
      const entry = getCompatibilityEntry(fixture.id)
      return entry?.wasmStatus === 'production' && !engineRunnableProductionFixtures.includes(fixture)
    })

    for (const fixture of skipped) {
      const documented = engineRuntimeSkipReasons.has(fixture.id) || runtimeJsOnlyFixtureIds.has(fixture.id)
      expect(documented, `Missing runtime skip reason for ${fixture.id}`).toBe(true)
      if (engineRuntimeSkipReasons.has(fixture.id)) {
        expect(engineRuntimeSkipReasons.get(fixture.id)).toMatch(/\S/)
      }
    }
  })

  it('keeps workbook metadata fixtures semantic on the JS route', async () => {
    await Promise.all(
      ['names:defined-name-range', 'names:defined-name-scalar', 'tables:table-total-row-sum', 'structured-reference:table-column-ref'].map(
        (fixtureId) => expectFixtureRuntimeResult(requiredCanonicalFixture(fixtureId), FormulaMode.JsOnly),
      ),
    )
  })

  it('keeps canonical grouped-array SUM fixtures in oracle parity on the wasm path', async () => {
    expect(groupedArrayProductionFixtures).toHaveLength(2)

    await Promise.all(
      groupedArrayProductionFixtures.map(async (fixture) => {
        try {
          await expectFixtureParity(fixture)
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          throw new Error(`Fixture ${fixture.id} failed: ${message}`, { cause: error })
        }
      }),
    )
  })

  it('keeps canonical criteria aggregate fixtures in oracle parity on the wasm path', async () => {
    expect(criteriaAggregateProductionFixtures).toHaveLength(criteriaAggregateProductionFixtureIds.size)

    await Promise.all(
      criteriaAggregateProductionFixtures.map(async (fixture) => {
        try {
          await expectFixtureParity(fixture)
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          throw new Error(`Fixture ${fixture.id} failed: ${message}`, { cause: error })
        }
      }),
    )
  })
})

async function expectFixtureParity(fixture: ExcelFixtureCase): Promise<void> {
  const { defaultSheetName, engine, ownerAddress, ownerSheetName } = await prepareFixtureEngine(fixture)
  evaluateBoundFixture({
    assertDifferential: true,
    defaultSheetName,
    engine,
    expectedMode: FormulaMode.WasmFastPath,
    fixture,
    ownerAddress,
    ownerSheetName,
  })
}

async function expectFixtureRuntimeResult(fixture: ExcelFixtureCase, expectedMode: FormulaMode): Promise<void> {
  const { defaultSheetName, engine, ownerAddress, ownerSheetName } = await prepareFixtureEngine(fixture)
  evaluateBoundFixture({
    assertDifferential: false,
    defaultSheetName,
    engine,
    expectedMode,
    fixture,
    ownerAddress,
    ownerSheetName,
  })
}

async function prepareFixtureEngine(fixture: ExcelFixtureCase): Promise<{
  readonly defaultSheetName: string
  readonly engine: SpreadsheetEngine
  readonly ownerAddress: string
  readonly ownerSheetName: string
}> {
  const engine = new SpreadsheetEngine({ workbookName: fixture.id })
  await engine.ready()

  const sheetNames = new Set<string>()
  const defaultSheetName = fixture.sheetName ?? 'Sheet1'
  sheetNames.add(defaultSheetName)

  fixture.inputs.forEach((input) => sheetNames.add(input.sheetName ?? defaultSheetName))
  fixture.outputs.forEach((output) => sheetNames.add(output.sheetName ?? defaultSheetName))

  for (const sheetName of sheetNames) {
    engine.createSheet(sheetName)
  }

  for (const input of fixture.inputs) {
    engine.setCellValue(input.sheetName ?? defaultSheetName, input.address, input.input)
  }
  applyFixtureMetadata(engine, fixture, defaultSheetName)

  const owner = fixture.outputs[0]
  if (!owner) {
    throw new Error(`Fixture ${fixture.id} is missing outputs`)
  }
  return {
    defaultSheetName,
    engine,
    ownerAddress: owner.address,
    ownerSheetName: owner.sheetName ?? defaultSheetName,
  }
}

function evaluateBoundFixture(args: {
  readonly assertDifferential: boolean
  readonly defaultSheetName: string
  readonly engine: SpreadsheetEngine
  readonly expectedMode: FormulaMode
  readonly fixture: ExcelFixtureCase
  readonly ownerAddress: string
  readonly ownerSheetName: string
}): void {
  withFixtureVolatileOracle(args.fixture, () => {
    args.engine.setCellFormula(args.ownerSheetName, args.ownerAddress, args.fixture.formula.replace(/^=/, ''))

    const explanation = args.engine.explainCell(args.ownerSheetName, args.ownerAddress)
    if (explanation.mode !== args.expectedMode) {
      throw new Error(`Fixture ${args.fixture.id} expected ${FormulaMode[args.expectedMode]} mode, received ${String(explanation.mode)}`)
    }

    if (args.assertDifferential) {
      const differential = args.engine.recalculateDifferential()
      if (differential.drift.length > 0) {
        throw new Error(`Fixture ${args.fixture.id} drifted between JS and wasm: ${JSON.stringify(differential.drift)}`)
      }
    }

    for (const output of args.fixture.outputs) {
      const actual = args.engine.getCellValue(output.sheetName ?? args.defaultSheetName, output.address)
      expectCellValueLike(actual, expectedValueToCellValue(output.expected))
    }
  })
}

function requiredCanonicalFixture(id: string): ExcelFixtureCase {
  const fixture = canonicalFormulaFixtures.find((candidate) => candidate.id === id)
  if (!fixture) {
    throw new Error(`Missing canonical formula fixture: ${id}`)
  }
  return fixture
}

function applyFixtureMetadata(engine: SpreadsheetEngine, fixture: ExcelFixtureCase, defaultSheetName: string): void {
  for (const definedName of fixture.definedNames ?? []) {
    engine.setDefinedName(definedName.name, fixtureDefinedNameValue(definedName.value))
  }
  for (const table of fixture.tables ?? []) {
    engine.setTable(fixtureTableSnapshot(table, defaultSheetName))
  }
}

function fixtureDefinedNameValue(value: WorkbookDefinedNameValueSnapshot): WorkbookDefinedNameValueSnapshot {
  return typeof value === 'string' && value.startsWith('=') ? { kind: 'formula', formula: value } : value
}

function fixtureTableSnapshot(table: ExcelFixtureTable, defaultSheetName: string): WorkbookTableSnapshot {
  return {
    name: table.name,
    sheetName: table.sheetName ?? defaultSheetName,
    startAddress: table.startAddress,
    endAddress: table.endAddress,
    columnNames: [...table.columnNames],
    headerRow: table.headerRow,
    totalsRow: table.totalsRow,
  }
}

function withFixtureVolatileOracle(fixture: ExcelFixtureCase, run: () => void): void {
  if (!capturedVolatileFixtureIds.has(fixture.id)) {
    run()
    return
  }
  vi.useFakeTimers()
  vi.setSystemTime(capturedVolatileOracleTime)
  const randomSpy = fixture.id === 'volatile:rand-basic' ? vi.spyOn(Math, 'random').mockReturnValue(0.625) : undefined
  try {
    run()
  } finally {
    randomSpy?.mockRestore()
    vi.useRealTimers()
  }
}

function expectedValueToCellValue(expected: ExcelExpectedValue): CellValue {
  switch (expected.kind) {
    case 'empty':
      return { tag: ValueTag.Empty }
    case 'number':
      return { tag: ValueTag.Number, value: expected.value }
    case 'boolean':
      return { tag: ValueTag.Boolean, value: expected.value }
    case 'string':
      return { tag: ValueTag.String, value: expected.value, stringId: 0 }
    case 'error':
      return { tag: ValueTag.Error, code: expected.code }
  }
}

function expectCellValueLike(actual: CellValue, expected: CellValue): void {
  expect(actual.tag).toBe(expected.tag)
  if (actual.tag === ValueTag.Number && expected.tag === ValueTag.Number) {
    expect(actual.value).toBeCloseTo(expected.value, 7)
    return
  }
  if (actual.tag === ValueTag.Error && expected.tag === ValueTag.Error) {
    expect(actual.code).toBe(expected.code)
    return
  }
  if (actual.tag === ValueTag.String && expected.tag === ValueTag.String) {
    expect(actual.value).toBe(expected.value)
    return
  }
  if (actual.tag === ValueTag.Boolean && expected.tag === ValueTag.Boolean) {
    expect(actual.value).toBe(expected.value)
    return
  }
  expect(actual).toEqual(expected)
}
