import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildCalculationSemanticsScorecard, parseCalculationSemanticsScorecard } from '../gen-calculation-semantics-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('calculation semantics scorecard', () => {
  it('covers every committed canonical and workbook-semantics fixture', () => {
    const scorecard = buildCalculationSemanticsScorecard()

    expect(scorecard.summary.allCommittedFormulaSemanticsCovered).toBe(true)
    expect(scorecard.summary.canonicalFormulaFixtureCount).toBe(301)
    expect(scorecard.summary.coveredCanonicalFixtureCount).toBe(301)
    expect(scorecard.summary.workbookSemanticsFixtureCount).toBe(12)
    expect(scorecard.summary.coveredWorkbookSemanticsFixtureCount).toBe(12)
    expect(scorecard.summary.coveredWorkbookSemanticsCategories).toEqual([
      'defined-names',
      'cross-sheet-references',
      'structured-references',
      'what-if-analysis',
      'dynamic-array-spills',
      'error-semantics',
    ])
    expect(scorecard.summary.missingCanonicalFixtureIds).toEqual([])
    expect(scorecard.summary.missingWorkbookSemanticsFixtureIds).toEqual([])
    expect(scorecard.summary.fixtureRegistryAligned).toBe(true)
    expect(scorecard.coverage.stableFormulaFixtureIds).toContain('lookup-reference:offset-basic')
    expect(scorecard.coverage.deterministicVolatileFixtureIds).toEqual([
      'date-time:now-volatile',
      'date-time:today-volatile',
      'volatile:rand-basic',
    ])
  })

  it('keeps the checked-in generated artifact aligned with the live fixture corpus', () => {
    const artifact = parseCalculationSemanticsScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/calculation-semantics-scorecard.json')),
    )
    const current = buildCalculationSemanticsScorecard(artifact.generatedAt)

    expect(artifact.summary).toEqual(current.summary)
    expect(artifact.coverage).toEqual(current.coverage)
  })
})
