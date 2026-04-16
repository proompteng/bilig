import type { WorkbookAgentExecutionPolicy } from '@bilig/contracts'
import type { WorkbookAgentBundleScope, WorkbookAgentCommandBundle, WorkbookAgentRiskClass } from './workbook-agent-bundles.js'

export interface WorkbookAgentExecutionPolicyInput {
  readonly scope: 'private' | 'shared'
  readonly executionPolicy: WorkbookAgentExecutionPolicy
  readonly riskClass: WorkbookAgentRiskClass
}

export type WorkbookAgentReviewDisposition = 'applyAutomatically' | 'reviewQueue'

export function resolveWorkbookAgentReviewDisposition(input: WorkbookAgentExecutionPolicyInput): WorkbookAgentReviewDisposition {
  if (
    requiresWorkbookAgentOwnerReview({
      scope: input.scope,
      riskClass: input.riskClass,
    }) ||
    input.executionPolicy === 'ownerReview'
  ) {
    return 'reviewQueue'
  }
  if (input.executionPolicy === 'autoApplyAll') {
    return 'applyAutomatically'
  }
  return input.riskClass === 'low' ? 'applyAutomatically' : 'reviewQueue'
}

export function isWorkbookAgentBundleAutoApplyEligible(input: WorkbookAgentExecutionPolicyInput): boolean {
  return resolveWorkbookAgentReviewDisposition(input) === 'applyAutomatically'
}

export function requiresWorkbookAgentOwnerReview(input: {
  readonly scope: 'private' | 'shared'
  readonly riskClass: WorkbookAgentRiskClass
}): boolean {
  return input.scope === 'shared' && input.riskClass !== 'low'
}

export function describeWorkbookAgentExecutionPolicy(policy: WorkbookAgentExecutionPolicy): string {
  switch (policy) {
    case 'autoApplySafe':
      return 'auto-apply safe changes'
    case 'autoApplyAll':
      return 'auto-apply all changes'
    case 'ownerReview':
      return 'owner review'
    default:
      return policy
  }
}

export function summarizeWorkbookAgentReviewTarget(input: {
  readonly summary: string
  readonly scope: WorkbookAgentBundleScope
  readonly riskClass: WorkbookAgentRiskClass
}): string {
  return `${input.summary} (${input.scope}, ${input.riskClass})`
}

export function resolveWorkbookAgentBundleExecutionPolicyInput(input: {
  readonly scope: 'private' | 'shared'
  readonly executionPolicy: WorkbookAgentExecutionPolicy
  readonly bundle: Pick<WorkbookAgentCommandBundle, 'riskClass'>
}): WorkbookAgentExecutionPolicyInput {
  return {
    scope: input.scope,
    executionPolicy: input.executionPolicy,
    riskClass: input.bundle.riskClass,
  }
}
