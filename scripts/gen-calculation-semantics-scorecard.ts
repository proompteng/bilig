#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  canonicalFormulaFixtures,
  canonicalWorkbookSemanticsFixtures,
  type ExcelFixtureCase,
  excelDateTimeFixtureSuite,
} from '../packages/excel-fixtures/src/index.ts'
import { formulaCompatibilityRegistry, getCompatibilityEntry } from '../packages/formula/src/compatibility.ts'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export interface WorkbookSemanticsCategoryCoverage {
  readonly category: string
  readonly fixtureIds: string[]
}

export interface CalculationSemanticsScorecard {
  readonly schemaVersion: 1
  readonly suite: 'calculation-semantics-coverage'
  readonly generatedAt: string
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-calculation-semantics-scorecard.ts'
    readonly fixturePackage: 'packages/excel-fixtures'
    readonly compatibilityRegistry: 'packages/formula/src/compatibility.ts'
    readonly fixtureHarnessTest: 'packages/formula/src/__tests__/fixture-harness.test.ts'
    readonly runtimeCorrectnessTest: 'packages/core/src/__tests__/formula-runtime-correctness.test.ts'
    readonly executionCommand: 'pnpm test:correctness:formula'
  }
  readonly summary: {
    readonly allCommittedFormulaSemanticsCovered: boolean
    readonly canonicalFormulaFixtureCount: number
    readonly workbookSemanticsFixtureCount: number
    readonly executableStableFormulaFixtureCount: number
    readonly deterministicVolatileFixtureCount: number
    readonly executableWorkbookSemanticsFixtureCount: number
    readonly dateTimeEdgeFixtureCount: number
    readonly coveredCanonicalFixtureCount: number
    readonly coveredWorkbookSemanticsFixtureCount: number
    readonly workbookSemanticsCategoryCount: number
    readonly coveredWorkbookSemanticsCategories: string[]
    readonly coveredFamilies: string[]
    readonly missingCanonicalFixtureIds: string[]
    readonly missingWorkbookSemanticsFixtureIds: string[]
    readonly fixtureRegistryAligned: boolean
  }
  readonly coverage: {
    readonly stableFormulaFixtureIds: string[]
    readonly deterministicVolatileFixtureIds: string[]
    readonly workbookSemanticsFixtureIds: string[]
    readonly workbookSemanticsCategories: WorkbookSemanticsCategoryCoverage[]
    readonly dateTimeEdgeFixtureIds: string[]
  }
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'calculation-semantics-scorecard.json')
const executableStatuses = new Set(['implemented-js', 'implemented-js-and-wasm-shadow', 'implemented-wasm-production'])
const deterministicVolatileFixtureIdSet = new Set(['date-time:today-volatile', 'date-time:now-volatile', 'volatile:rand-basic'])
const workbookSemanticsCategorySpecs = [
  {
    category: 'defined-names',
    matches: (fixture: ExcelFixtureCase) => (fixture.definedNames?.length ?? 0) > 0,
  },
  {
    category: 'cross-sheet-references',
    matches: (fixture: ExcelFixtureCase) =>
      fixture.formula.includes('!') ||
      fixture.inputs.some((input) => input.sheetName !== undefined && input.sheetName !== fixture.sheetName),
  },
  {
    category: 'structured-references',
    matches: (fixture: ExcelFixtureCase) => (fixture.tables?.length ?? 0) > 0 || /\[[^\]]+\]/u.test(fixture.formula),
  },
  {
    category: 'what-if-analysis',
    matches: (fixture: ExcelFixtureCase) => fixture.multipleOperations !== undefined,
  },
  {
    category: 'dynamic-array-spills',
    matches: (fixture: ExcelFixtureCase) => fixture.outputs.length > 1,
  },
  {
    category: 'error-semantics',
    matches: (fixture: ExcelFixtureCase) => fixture.outputs.some((output) => output.expected.kind === 'error'),
  },
] as const

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  const scorecard = buildCalculationSemanticsScorecard()
  const serializedScorecard = formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`)

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(`Calculation semantics scorecard is missing. Run: bun scripts/gen-calculation-semantics-scorecard.ts`)
    }
    const currentScorecard = readFileSync(outputPath, 'utf8')
    if (currentScorecard !== serializedScorecard) {
      throw new Error('Generated calculation semantics scorecard is out of date. Run: bun scripts/gen-calculation-semantics-scorecard.ts')
    }
    validateCalculationSemanticsScorecard(parseCalculationSemanticsScorecard(readJsonObject(outputPath)))
  } else {
    mkdirSync(dirname(outputPath), { recursive: true })
    writeFileSync(outputPath, serializedScorecard)
  }

  console.log(
    JSON.stringify(
      {
        mode: isCheckMode ? 'check' : 'write',
        outputPath,
        allCommittedFormulaSemanticsCovered: scorecard.summary.allCommittedFormulaSemanticsCovered,
        coveredCanonicalFixtureCount: scorecard.summary.coveredCanonicalFixtureCount,
        canonicalFormulaFixtureCount: scorecard.summary.canonicalFormulaFixtureCount,
        coveredWorkbookSemanticsFixtureCount: scorecard.summary.coveredWorkbookSemanticsFixtureCount,
        workbookSemanticsFixtureCount: scorecard.summary.workbookSemanticsFixtureCount,
      },
      null,
      2,
    ),
  )
}

export function buildCalculationSemanticsScorecard(generatedAt = 'checked-in-generated-artifact'): CalculationSemanticsScorecard {
  const stableFormulaFixtureIds = canonicalFormulaFixtures
    .filter(isStableExecutableFormulaFixture)
    .map((fixture) => fixture.id)
    .toSorted()
  const deterministicVolatileFixtureIds = canonicalFormulaFixtures
    .filter((fixture) => deterministicVolatileFixtureIdSet.has(fixture.id))
    .map((fixture) => fixture.id)
    .toSorted()
  const workbookSemanticsFixtureIds = canonicalWorkbookSemanticsFixtures
    .filter((fixture) => executableStatuses.has(getCompatibilityEntry(fixture.id)?.status ?? 'unsupported'))
    .map((fixture) => fixture.id)
    .toSorted()
  const workbookSemanticsCategories = buildWorkbookSemanticsCategoryCoverage(canonicalWorkbookSemanticsFixtures)
  const coveredWorkbookSemanticsCategories = workbookSemanticsCategories.map((entry) => entry.category)
  const dateTimeEdgeFixtureIds = (excelDateTimeFixtureSuite.cases ?? [])
    .filter(isStableExecutableFormulaFixture)
    .map((fixture) => fixture.id)
    .toSorted()
  const coveredCanonicalFixtureIds = new Set([...stableFormulaFixtureIds, ...deterministicVolatileFixtureIds])
  const missingCanonicalFixtureIds = canonicalFormulaFixtures
    .map((fixture) => fixture.id)
    .filter((fixtureId) => !coveredCanonicalFixtureIds.has(fixtureId))
    .toSorted()
  const missingWorkbookSemanticsFixtureIds = canonicalWorkbookSemanticsFixtures
    .map((fixture) => fixture.id)
    .filter((fixtureId) => !workbookSemanticsFixtureIds.includes(fixtureId))
    .toSorted()
  const canonicalRegistryFixtureIds = formulaCompatibilityRegistry
    .filter((entry) => entry.scope === 'canonical')
    .map((entry) => entry.id)
    .toSorted()
  const canonicalFixtureIds = canonicalFormulaFixtures.map((fixture) => fixture.id).toSorted()
  const coveredFamilies = [
    ...new Set([...canonicalFormulaFixtures, ...canonicalWorkbookSemanticsFixtures].map((fixture) => fixture.family)),
  ].toSorted()
  const fixtureRegistryAligned = JSON.stringify(canonicalRegistryFixtureIds) === JSON.stringify(canonicalFixtureIds)

  return {
    schemaVersion: 1,
    suite: 'calculation-semantics-coverage',
    generatedAt,
    source: {
      artifactGenerator: 'scripts/gen-calculation-semantics-scorecard.ts',
      fixturePackage: 'packages/excel-fixtures',
      compatibilityRegistry: 'packages/formula/src/compatibility.ts',
      fixtureHarnessTest: 'packages/formula/src/__tests__/fixture-harness.test.ts',
      runtimeCorrectnessTest: 'packages/core/src/__tests__/formula-runtime-correctness.test.ts',
      executionCommand: 'pnpm test:correctness:formula',
    },
    summary: {
      allCommittedFormulaSemanticsCovered:
        fixtureRegistryAligned && missingCanonicalFixtureIds.length === 0 && missingWorkbookSemanticsFixtureIds.length === 0,
      canonicalFormulaFixtureCount: canonicalFormulaFixtures.length,
      workbookSemanticsFixtureCount: canonicalWorkbookSemanticsFixtures.length,
      executableStableFormulaFixtureCount: stableFormulaFixtureIds.length,
      deterministicVolatileFixtureCount: deterministicVolatileFixtureIds.length,
      executableWorkbookSemanticsFixtureCount: workbookSemanticsFixtureIds.length,
      dateTimeEdgeFixtureCount: dateTimeEdgeFixtureIds.length,
      coveredCanonicalFixtureCount: coveredCanonicalFixtureIds.size,
      coveredWorkbookSemanticsFixtureCount: workbookSemanticsFixtureIds.length,
      workbookSemanticsCategoryCount: workbookSemanticsCategories.length,
      coveredWorkbookSemanticsCategories,
      coveredFamilies,
      missingCanonicalFixtureIds,
      missingWorkbookSemanticsFixtureIds,
      fixtureRegistryAligned,
    },
    coverage: {
      stableFormulaFixtureIds,
      deterministicVolatileFixtureIds,
      workbookSemanticsFixtureIds,
      workbookSemanticsCategories,
      dateTimeEdgeFixtureIds,
    },
  }
}

export function parseCalculationSemanticsScorecard(value: Record<string, unknown>): CalculationSemanticsScorecard {
  const source = asObject(value['source'], 'calculation semantics source')
  const summary = asObject(value['summary'], 'calculation semantics summary')
  const coverage = asObject(value['coverage'], 'calculation semantics coverage')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'calculation-semantics-coverage'),
    generatedAt: typeof value['generatedAt'] === 'string' ? value['generatedAt'] : '',
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-calculation-semantics-scorecard.ts'),
      fixturePackage: literalField(source, 'fixturePackage', 'packages/excel-fixtures'),
      compatibilityRegistry: literalField(source, 'compatibilityRegistry', 'packages/formula/src/compatibility.ts'),
      fixtureHarnessTest: literalField(source, 'fixtureHarnessTest', 'packages/formula/src/__tests__/fixture-harness.test.ts'),
      runtimeCorrectnessTest: literalField(
        source,
        'runtimeCorrectnessTest',
        'packages/core/src/__tests__/formula-runtime-correctness.test.ts',
      ),
      executionCommand: literalField(source, 'executionCommand', 'pnpm test:correctness:formula'),
    },
    summary: {
      allCommittedFormulaSemanticsCovered: booleanField(summary, 'allCommittedFormulaSemanticsCovered'),
      canonicalFormulaFixtureCount: numberField(summary, 'canonicalFormulaFixtureCount'),
      workbookSemanticsFixtureCount: numberField(summary, 'workbookSemanticsFixtureCount'),
      executableStableFormulaFixtureCount: numberField(summary, 'executableStableFormulaFixtureCount'),
      deterministicVolatileFixtureCount: numberField(summary, 'deterministicVolatileFixtureCount'),
      executableWorkbookSemanticsFixtureCount: numberField(summary, 'executableWorkbookSemanticsFixtureCount'),
      dateTimeEdgeFixtureCount: numberField(summary, 'dateTimeEdgeFixtureCount'),
      coveredCanonicalFixtureCount: numberField(summary, 'coveredCanonicalFixtureCount'),
      coveredWorkbookSemanticsFixtureCount: numberField(summary, 'coveredWorkbookSemanticsFixtureCount'),
      workbookSemanticsCategoryCount: numberField(summary, 'workbookSemanticsCategoryCount'),
      coveredWorkbookSemanticsCategories: stringArrayField(summary, 'coveredWorkbookSemanticsCategories'),
      coveredFamilies: stringArrayField(summary, 'coveredFamilies'),
      missingCanonicalFixtureIds: stringArrayField(summary, 'missingCanonicalFixtureIds'),
      missingWorkbookSemanticsFixtureIds: stringArrayField(summary, 'missingWorkbookSemanticsFixtureIds'),
      fixtureRegistryAligned: booleanField(summary, 'fixtureRegistryAligned'),
    },
    coverage: {
      stableFormulaFixtureIds: stringArrayField(coverage, 'stableFormulaFixtureIds'),
      deterministicVolatileFixtureIds: stringArrayField(coverage, 'deterministicVolatileFixtureIds'),
      workbookSemanticsFixtureIds: stringArrayField(coverage, 'workbookSemanticsFixtureIds'),
      workbookSemanticsCategories: workbookSemanticsCategoryCoverageArrayField(coverage, 'workbookSemanticsCategories'),
      dateTimeEdgeFixtureIds: stringArrayField(coverage, 'dateTimeEdgeFixtureIds'),
    },
  }
}

export function validateCalculationSemanticsScorecard(scorecard: CalculationSemanticsScorecard): void {
  const current = buildCalculationSemanticsScorecard(scorecard.generatedAt)
  if (
    JSON.stringify(scorecard.summary) !== JSON.stringify(current.summary) ||
    JSON.stringify(scorecard.coverage) !== JSON.stringify(current.coverage)
  ) {
    throw new Error('Calculation semantics scorecard is stale against the current fixture corpus')
  }
  if (!scorecard.summary.allCommittedFormulaSemanticsCovered) {
    throw new Error(
      `Calculation semantics coverage is incomplete: missing canonical=${scorecard.summary.missingCanonicalFixtureIds.join(
        ',',
      )}; missing workbook=${scorecard.summary.missingWorkbookSemanticsFixtureIds.join(',')}`,
    )
  }
  if (!scorecard.coverage.stableFormulaFixtureIds.includes('lookup-reference:offset-basic')) {
    throw new Error('Calculation semantics scorecard must cover the canonical OFFSET fixture')
  }
  if (!scorecard.summary.coveredWorkbookSemanticsCategories.includes('structured-references')) {
    throw new Error('Calculation semantics scorecard must cover structured-reference workbook semantics')
  }
}

function buildWorkbookSemanticsCategoryCoverage(fixtures: readonly ExcelFixtureCase[]): WorkbookSemanticsCategoryCoverage[] {
  return workbookSemanticsCategorySpecs
    .map((spec) => ({
      category: spec.category,
      fixtureIds: fixtures
        .filter(spec.matches)
        .map((fixture) => fixture.id)
        .toSorted(),
    }))
    .filter((entry) => entry.fixtureIds.length > 0)
}

function workbookSemanticsCategoryCoverageArrayField(value: Record<string, unknown>, field: string): WorkbookSemanticsCategoryCoverage[] {
  return arrayField(value, field).map((entry, index) => {
    const category = asObject(entry, `${field}[${String(index)}]`)
    return {
      category: stringField(category, 'category'),
      fixtureIds: stringArrayField(category, 'fixtureIds'),
    }
  })
}

function isStableExecutableFormulaFixture(fixture: { readonly family: string; readonly formula: string; readonly id: string }): boolean {
  const entry = getCompatibilityEntry(fixture.id)
  const hasVolatileCall = /\b(TODAY|NOW|RAND)\s*\(/iu.test(fixture.formula)
  return entry !== undefined && executableStatuses.has(entry.status) && fixture.family !== 'volatile' && !hasVolatileCall
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
