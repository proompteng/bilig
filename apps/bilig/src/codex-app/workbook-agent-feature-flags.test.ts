import { describe, expect, it } from 'vitest'
import {
  getWorkbookAgentWorkflowFamily,
  isWorkbookAgentRolloutAllowed,
  isWorkbookAgentWorkflowFamilyEnabled,
  resolveWorkbookAgentFeatureFlags,
} from './workbook-agent-feature-flags.js'

describe('workbook agent feature flags', () => {
  it('fails closed outside local environments unless the agent is explicitly enabled', () => {
    expect(resolveWorkbookAgentFeatureFlags({ NODE_ENV: 'production' } as NodeJS.ProcessEnv).enabled).toBe(false)
    expect(resolveWorkbookAgentFeatureFlags({ NODE_ENV: 'staging' } as NodeJS.ProcessEnv).enabled).toBe(false)
    expect(resolveWorkbookAgentFeatureFlags({} as NodeJS.ProcessEnv).enabled).toBe(false)
    expect(resolveWorkbookAgentFeatureFlags({ NODE_ENV: 'development' } as NodeJS.ProcessEnv).enabled).toBe(true)
    expect(resolveWorkbookAgentFeatureFlags({ NODE_ENV: 'test' } as NodeJS.ProcessEnv).enabled).toBe(true)
    expect(
      resolveWorkbookAgentFeatureFlags({
        NODE_ENV: 'production',
        BILIG_AGENT_ENABLED: 'true',
      } as NodeJS.ProcessEnv).enabled,
    ).toBe(true)
  })

  it('resolves feature flags from environment values', () => {
    expect(
      resolveWorkbookAgentFeatureFlags({
        BILIG_AGENT_ENABLED: 'false',
        BILIG_AGENT_SHARED_THREADS_ENABLED: 'false',
        BILIG_AGENT_WORKFLOW_RUNNER_ENABLED: '1',
        BILIG_AGENT_AUTO_APPLY_LOW_RISK_ENABLED: '0',
        BILIG_AGENT_FORMULA_WORKFLOWS_ENABLED: 'false',
        BILIG_AGENT_FORMATTING_WORKFLOWS_ENABLED: 'true',
        BILIG_AGENT_IMPORT_WORKFLOWS_ENABLED: 'true',
        BILIG_AGENT_ROLLUP_WORKFLOWS_ENABLED: 'false',
        BILIG_AGENT_STRUCTURAL_WORKFLOWS_ENABLED: 'true',
        BILIG_AGENT_ALLOWLIST_USERS: 'alex@example.com, pat@example.com',
        BILIG_AGENT_ALLOWLIST_DOCUMENTS: 'doc-1',
      } as NodeJS.ProcessEnv),
    ).toEqual(
      expect.objectContaining({
        enabled: false,
        sharedThreadsEnabled: false,
        workflowRunnerEnabled: true,
        autoApplyLowRiskEnabled: false,
        formulaWorkflowFamilyEnabled: false,
        rollupWorkflowFamilyEnabled: false,
        allowlistedUserIds: ['alex@example.com', 'pat@example.com'],
        allowlistedDocumentIds: ['doc-1'],
      }),
    )
  })

  it('rejects malformed boolean feature flags instead of silently using defaults', () => {
    expect(() =>
      resolveWorkbookAgentFeatureFlags({
        BILIG_AGENT_ENABLED: 'sometimes',
      } as NodeJS.ProcessEnv),
    ).toThrow('BILIG_AGENT_ENABLED must be "1", "true", "0", or "false" when set, got sometimes')

    expect(() =>
      resolveWorkbookAgentFeatureFlags({
        BILIG_AGENT_AUTO_APPLY_LOW_RISK_ENABLED: 'yes',
      } as NodeJS.ProcessEnv),
    ).toThrow('BILIG_AGENT_AUTO_APPLY_LOW_RISK_ENABLED must be "1", "true", "0", or "false" when set, got yes')

    expect(() =>
      resolveWorkbookAgentFeatureFlags({
        BILIG_AGENT_WORKFLOW_RUNNER_ENABLED: 'off',
      } as NodeJS.ProcessEnv),
    ).toThrow('BILIG_AGENT_WORKFLOW_RUNNER_ENABLED must be "1", "true", "0", or "false" when set, got off')
  })

  it('checks rollout allowlists by user or document', () => {
    expect(
      isWorkbookAgentRolloutAllowed(
        {
          allowlistedUserIds: ['alex@example.com'],
          allowlistedDocumentIds: ['doc-2'],
        },
        { documentId: 'doc-1', userId: 'alex@example.com' },
      ),
    ).toBe(true)
    expect(
      isWorkbookAgentRolloutAllowed(
        {
          allowlistedUserIds: ['alex@example.com'],
          allowlistedDocumentIds: ['doc-2'],
        },
        { documentId: 'doc-2', userId: 'pat@example.com' },
      ),
    ).toBe(true)
    expect(
      isWorkbookAgentRolloutAllowed(
        {
          allowlistedUserIds: ['alex@example.com'],
          allowlistedDocumentIds: ['doc-2'],
        },
        { documentId: 'doc-1', userId: 'pat@example.com' },
      ),
    ).toBe(false)
  })

  it('maps workflow templates to families and enablement flags', () => {
    const featureFlags = {
      formulaWorkflowFamilyEnabled: false,
      formattingWorkflowFamilyEnabled: true,
      importWorkflowFamilyEnabled: false,
      rollupWorkflowFamilyEnabled: true,
      structuralWorkflowFamilyEnabled: false,
    }

    expect(getWorkbookAgentWorkflowFamily('summarizeWorkbook')).toBe('report')
    expect(getWorkbookAgentWorkflowFamily('highlightFormulaIssues')).toBe('formula')
    expect(getWorkbookAgentWorkflowFamily('styleCurrentSheetHeaders')).toBe('formatting')
    expect(getWorkbookAgentWorkflowFamily('normalizeCurrentSheetHeaders')).toBe('import')
    expect(getWorkbookAgentWorkflowFamily('createCurrentSheetRollup')).toBe('rollup')
    expect(getWorkbookAgentWorkflowFamily('createSheet')).toBe('structural')

    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'summarizeWorkbook')).toBe(true)
    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'highlightFormulaIssues')).toBe(false)
    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'styleCurrentSheetHeaders')).toBe(true)
    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'normalizeCurrentSheetHeaders')).toBe(false)
    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'createCurrentSheetRollup')).toBe(true)
    expect(isWorkbookAgentWorkflowFamilyEnabled(featureFlags, 'createSheet')).toBe(false)
  })
})
