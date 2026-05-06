import { describe, expect, it } from 'vitest'

import { buildSecurityPostureScorecard, validateSecurityPostureScorecard } from '../gen-security-posture-scorecard.ts'

describe('security posture scorecard', () => {
  it('generates a checked artifact from executable security posture controls', () => {
    const scorecard = buildSecurityPostureScorecard('2026-05-06T09:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'security-posture',
      generatedAt: '2026-05-06T09:00:00.000Z',
      summary: {
        allRequiredControlsPassed: true,
        formulaSandboxPassed: true,
        importSafetyPassed: true,
        agentPermissionPolicyPassed: true,
        runtimePackageHardeningPassed: true,
        browserCspPassed: true,
        dependencyAuditPassed: true,
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
      },
    })
    expect(scorecard.controls.map((control) => control.id)).toEqual([
      'formula-runtime-no-dynamic-code-execution',
      'xlsx-import-macro-non-execution',
      'shared-agent-owner-review',
      'runtime-publish-package-hardening',
      'browser-content-security-policy',
      'production-dependency-vulnerability-audit',
    ])
    expect(scorecard.controls.every((control) => control.required && control.passed)).toBe(true)
    expect(scorecard.summary.coveredControls).toEqual([
      'formula.noEval',
      'formula.noFunctionConstructor',
      'formula.noNodeProcessExecution',
      'xlsx.macroWarning',
      'xlsx.noMacroPayloadExport',
      'agent.sharedMediumHighRiskOwnerReview',
      'runtime.publishManifest',
      'runtime.noSourceInTarballs',
      'runtime.alignedPackageSet',
      'browser.contentSecurityPolicy',
      'browser.crossOriginIsolation',
      'browser.workerWasmRuntimeAllowlist',
      'dependency.vulnerabilityAudit',
    ])
    expect(scorecard.summary.uncoveredControls).toEqual(['deployment.runtimeNetworkPolicy', 'externalSheetsExcelSecurityComparison'])
  })

  it('rejects stale artifacts missing required security controls', () => {
    const scorecard = buildSecurityPostureScorecard('2026-05-06T09:00:00.000Z')
    const staleScorecard = {
      ...scorecard,
      controls: scorecard.controls.filter((control) => control.id !== 'shared-agent-owner-review'),
    }

    expect(() => validateSecurityPostureScorecard(staleScorecard)).toThrow(
      'Security posture scorecard is missing required control: shared-agent-owner-review',
    )
  })
})
