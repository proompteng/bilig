import { describe, expect, it } from 'vitest'

import { buildAutomationScorecard, validateAutomationScorecard } from '../gen-automation-scorecard.ts'

describe('automation scorecard', () => {
  it('executes real semantic workbook automation controls', async () => {
    const scorecard = await buildAutomationScorecard('2026-05-06T12:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'automation-api-extensibility',
      generatedAt: '2026-05-06T12:00:00.000Z',
      summary: {
        allRequiredControlsPassed: true,
        semanticCommandWorkflowPassed: true,
        headlessServiceWorkflowPassed: true,
        workerPreviewWorkflowPassed: true,
        toolRegistryPassed: true,
        tenXWorkflowAutomationBenchmarkPassed: true,
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
      },
    })
    expect(scorecard.controls.map((control) => control.id)).toEqual([
      'semantic-command-bundle-preview-apply',
      'headless-service-automation-workflow',
      'worker-runtime-agent-preview',
      'agent-tool-registry-semantic-coverage',
      'semantic-workflow-automation-ten-x-benchmark',
    ])
    expect(scorecard.controls.every((control) => control.required && control.passed)).toBe(true)
    expect(scorecard.summary.coveredControls).toEqual([
      'agent.semanticBundleValidation',
      'agent.previewApplyExecution',
      'agent.partialCommandProjection',
      'headless.serviceWorkflow',
      'headless.persistenceRoundTrip',
      'headless.undoRedoAutomation',
      'worker.runtimePreview',
      'tools.semanticWorkbookRegistry',
      'tools.legacyNameNormalization',
      'automation.tenXWorkflowBenchmark',
    ])
    expect(scorecard.summary.uncoveredControls).toEqual(['googleAppsScriptDirectComparison', 'officeScriptsDirectComparison'])

    validateAutomationScorecard(scorecard)
  })

  it('fails validation when a required automation control is missing', async () => {
    const scorecard = await buildAutomationScorecard('2026-05-06T12:00:00.000Z')

    expect(() =>
      validateAutomationScorecard({
        ...scorecard,
        controls: scorecard.controls.filter((control) => control.id !== 'worker-runtime-agent-preview'),
      }),
    ).toThrow('Automation scorecard is missing required control: worker-runtime-agent-preview')
  })
})
