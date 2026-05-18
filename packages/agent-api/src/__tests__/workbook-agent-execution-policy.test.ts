import { describe, expect, it } from 'vitest'
import {
  canCancelWorkbookAgentWorkflowRun,
  canInterruptWorkbookAgentTurn,
  describeWorkbookAgentExecutionPolicy,
  isWorkbookAgentBundleAutoApplyEligible,
  requiresWorkbookAgentOwnerReview,
  resolveWorkbookAgentReviewDisposition,
} from '../workbook-agent-execution-policy.js'

describe('workbook agent execution policy', () => {
  it('routes private auto-apply-all sessions straight to execution', () => {
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: 'private',
        executionPolicy: 'autoApplyAll',
        riskClass: 'high',
      }),
    ).toBe(true)
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: 'private',
        executionPolicy: 'autoApplyAll',
        riskClass: 'high',
      }),
    ).toBe('applyAutomatically')
  })

  it('routes private safe sessions by risk class', () => {
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: 'private',
        executionPolicy: 'autoApplySafe',
        riskClass: 'low',
      }),
    ).toBe(true)
    expect(
      isWorkbookAgentBundleAutoApplyEligible({
        scope: 'private',
        executionPolicy: 'autoApplySafe',
        riskClass: 'medium',
      }),
    ).toBe(false)
  })

  it('routes shared medium and high risk work through owner review', () => {
    expect(
      requiresWorkbookAgentOwnerReview({
        scope: 'shared',
        riskClass: 'medium',
      }),
    ).toBe(true)
    expect(
      requiresWorkbookAgentOwnerReview({
        scope: 'shared',
        riskClass: 'low',
      }),
    ).toBe(false)
  })

  it('allows shared low-risk work to execute directly when session policy allows it', () => {
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: 'shared',
        executionPolicy: 'autoApplyAll',
        riskClass: 'low',
      }),
    ).toBe('applyAutomatically')
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: 'shared',
        executionPolicy: 'autoApplySafe',
        riskClass: 'low',
      }),
    ).toBe('applyAutomatically')
  })

  it('keeps shared medium and high risk work behind review even under auto-apply policies', () => {
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: 'shared',
        executionPolicy: 'autoApplyAll',
        riskClass: 'medium',
      }),
    ).toBe('reviewQueue')
    expect(
      resolveWorkbookAgentReviewDisposition({
        scope: 'shared',
        executionPolicy: 'autoApplySafe',
        riskClass: 'high',
      }),
    ).toBe('reviewQueue')
  })

  it('describes execution policies for user-facing summaries', () => {
    expect(describeWorkbookAgentExecutionPolicy('autoApplySafe')).toBe('auto-apply safe changes')
    expect(describeWorkbookAgentExecutionPolicy('autoApplyAll')).toBe('auto-apply all changes')
    expect(describeWorkbookAgentExecutionPolicy('ownerReview')).toBe('owner review')
  })

  it('limits shared workflow cancellation to the starter or thread owner', () => {
    expect(
      canCancelWorkbookAgentWorkflowRun({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'casey@example.com',
        startedByUserId: 'casey@example.com',
      }),
    ).toBe(true)
    expect(
      canCancelWorkbookAgentWorkflowRun({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'alex@example.com',
        startedByUserId: 'casey@example.com',
      }),
    ).toBe(true)
    expect(
      canCancelWorkbookAgentWorkflowRun({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'pat@example.com',
        startedByUserId: 'casey@example.com',
      }),
    ).toBe(false)
  })

  it('limits shared turn interruption to the active turn actor or thread owner', () => {
    expect(
      canInterruptWorkbookAgentTurn({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'casey@example.com',
        turnActorUserId: 'casey@example.com',
      }),
    ).toBe(true)
    expect(
      canInterruptWorkbookAgentTurn({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'alex@example.com',
        turnActorUserId: 'casey@example.com',
      }),
    ).toBe(true)
    expect(
      canInterruptWorkbookAgentTurn({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'pat@example.com',
        turnActorUserId: 'casey@example.com',
      }),
    ).toBe(false)
    expect(
      canInterruptWorkbookAgentTurn({
        scope: 'shared',
        ownerUserId: 'alex@example.com',
        actorUserId: 'pat@example.com',
        turnActorUserId: null,
      }),
    ).toBe(false)
  })
})
