import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { summarizeWorkbookAgentVerificationStatus } from '../workbook-agent-verification-status.js'

describe('workbook agent verification status fuzz', () => {
  it('should make verification completeness exactly match required generated checks', async () => {
    await runProperty({
      suite: 'bilig/codex-app/verification-status/completeness',
      arbitrary: fc.record({
        renderedReadback: fc.array(
          fc.record({
            requested: fc.boolean(),
            matched: fc.option(fc.boolean(), { nil: null }),
          }),
          { maxLength: 4 },
        ),
        formulaIssueCount: fc.integer({ min: 0, max: 3 }),
        invariantsOk: fc.boolean(),
        includeFormulaReport: fc.boolean(),
        includeInvariantReport: fc.boolean(),
        requireTargetRange: fc.boolean(),
        targetRangeCount: fc.integer({ min: 0, max: 3 }),
      }),
      predicate: async (input) => {
        const status = summarizeWorkbookAgentVerificationStatus({
          renderedReadback: input.renderedReadback,
          formulaIssues: input.includeFormulaReport
            ? {
                summary: { actionableIssueCount: input.formulaIssueCount },
              }
            : null,
          invariants: input.includeInvariantReport
            ? {
                summary: { ok: input.invariantsOk },
              }
            : null,
          requireTargetRange: input.requireTargetRange,
          targetRangeCount: input.targetRangeCount,
        })

        const expectedRendered = input.renderedReadback.every((proof) => !proof.requested || proof.matched === true)
        const expectedFormula = input.includeFormulaReport && input.formulaIssueCount === 0
        const expectedInvariants = input.includeInvariantReport && input.invariantsOk
        const expectedTarget = input.requireTargetRange ? input.targetRangeCount > 0 : true

        expect(status.renderedComplete).toBe(expectedRendered)
        expect(status.formulaComplete).toBe(expectedFormula)
        expect(status.invariantsComplete).toBe(expectedInvariants)
        expect(status.targetRangeComplete).toBe(expectedTarget)
        expect(status.verificationComplete).toBe(expectedRendered && expectedFormula && expectedInvariants && expectedTarget)
      },
      parameters: { numRuns: 120 },
    })
  })
})
