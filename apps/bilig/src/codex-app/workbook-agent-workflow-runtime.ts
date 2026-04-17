import { createWorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookAgentExecutionRecord } from '@bilig/agent-api'
import type { WorkbookAgentWorkflowRun } from '@bilig/contracts'
import type { ZeroSyncService } from '../zero/service.js'
import { attachSharedReviewState } from './workbook-agent-bundle-state.js'
import {
  cancelWorkflowSteps,
  completeWorkflowSteps,
  createWorkflowRunRecord,
  createRunningWorkflowSteps,
  describeWorkbookAgentWorkflowTemplate,
  executeWorkbookAgentWorkflow,
  failWorkflowSteps,
} from './workbook-agent-workflows.js'
import { isWorkflowAbortError, throwIfWorkflowCancelled } from './workbook-agent-workflow-abort.js'
import { createSystemEntry } from './workbook-agent-session-model.js'
import {
  type QueuedWorkbookAgentWorkflowRun,
  type WorkbookAgentThreadState,
  type WorkbookAgentWorkflowInput,
  toContextRef,
  upsertEntry,
  upsertWorkflowRun,
} from './workbook-agent-service-shared.js'

export class WorkbookAgentWorkflowRuntime {
  private readonly workflowRunTasks = new Map<string, Promise<void>>()
  private readonly workflowAbortControllers = new Map<string, AbortController>()

  constructor(
    private readonly options: {
      zeroSyncService: ZeroSyncService
      now: () => number
      touch: (sessionState: WorkbookAgentThreadState) => void
      persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>
      emitSnapshot: (threadId: string) => void
      shouldApplyBundleImmediately: (
        sessionState: WorkbookAgentThreadState,
        bundle: ReturnType<typeof createWorkbookAgentCommandBundle>,
      ) => boolean
      stageReviewBundle: (
        sessionState: WorkbookAgentThreadState,
        turnId: string,
        bundle: ReturnType<typeof createWorkbookAgentCommandBundle>,
      ) => void
      applyCommandBundleAutomatically: (input: {
        sessionState: WorkbookAgentThreadState
        actorUserId: string
        bundle: ReturnType<typeof createWorkbookAgentCommandBundle>
      }) => Promise<WorkbookAgentExecutionRecord | null>
      incrementCounter: (
        counter: 'workflowStartedCount' | 'workflowCompletedCount' | 'workflowFailedCount' | 'workflowCancelledCount',
      ) => void
    },
  ) {}

  async startWorkflow(input: {
    sessionState: WorkbookAgentThreadState
    documentId: string
    actorUserId: string
    workflowTemplate: WorkbookAgentWorkflowRun['workflowTemplate']
    workflowInput: WorkbookAgentWorkflowInput
    workflowTurnId: string
  }): Promise<void> {
    const runId = crypto.randomUUID()
    const now = this.options.now()
    const workflowDescription = describeWorkbookAgentWorkflowTemplate(input.workflowTemplate, input.workflowInput)
    const runningRun = createWorkflowRunRecord({
      runId,
      threadId: input.sessionState.threadId,
      startedByUserId: input.actorUserId,
      workflowTemplate: input.workflowTemplate,
      title: workflowDescription.title,
      summary: workflowDescription.runningSummary,
      status: 'running',
      now,
      steps: createRunningWorkflowSteps(input.workflowTemplate, now, input.workflowInput),
    })
    input.sessionState.durable.workflowRuns = upsertWorkflowRun(input.sessionState.durable.workflowRuns, runningRun)
    input.sessionState.durable.entries = upsertEntry(
      input.sessionState.durable.entries,
      createSystemEntry(`system-workflow-start:${runId}`, input.workflowTurnId, `Started workflow: ${runningRun.title}`),
    )
    this.options.touch(input.sessionState)
    this.options.incrementCounter('workflowStartedCount')
    await this.options.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, runningRun)
    await this.options.persistSessionState(input.sessionState)
    this.options.emitSnapshot(input.sessionState.threadId)
    this.queueWorkflowRun({
      sessionState: input.sessionState,
      documentId: input.documentId,
      runId,
      workflowTurnId: input.workflowTurnId,
      workflowTemplate: input.workflowTemplate,
      workflowInput: input.workflowInput,
      startedByUserId: input.actorUserId,
      runningRun,
    })
  }

  async cancelRunningWorkflow(input: {
    sessionState: WorkbookAgentThreadState
    documentId: string
    runId: string
    runningWorkflow: WorkbookAgentWorkflowRun
    actorUserId: string
  }): Promise<void> {
    const now = this.options.now()
    const cancelledRun: WorkbookAgentWorkflowRun = {
      ...input.runningWorkflow,
      status: 'cancelled',
      summary: `Cancelled workflow: ${input.runningWorkflow.title}`,
      updatedAtUnixMs: now,
      completedAtUnixMs: now,
      errorMessage: `Cancelled by ${input.actorUserId}.`,
      steps: cancelWorkflowSteps(input.runningWorkflow.steps, now),
      artifact: null,
    }
    input.sessionState.durable.workflowRuns = upsertWorkflowRun(input.sessionState.durable.workflowRuns, cancelledRun)
    input.sessionState.durable.entries = upsertEntry(
      input.sessionState.durable.entries,
      createSystemEntry(
        `system-workflow-cancel:${input.runId}:${now}`,
        input.sessionState.live.activeTurnId,
        `Cancelled workflow: ${input.runningWorkflow.title}`,
      ),
    )
    this.workflowAbortControllers.get(input.runId)?.abort()
    this.options.touch(input.sessionState)
    this.options.incrementCounter('workflowCancelledCount')
    await this.options.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, cancelledRun)
    await this.options.persistSessionState(input.sessionState)
    this.options.emitSnapshot(input.sessionState.threadId)
  }

  close(): void {
    this.workflowAbortControllers.forEach((controller) => {
      controller.abort()
    })
    this.workflowAbortControllers.clear()
    this.workflowRunTasks.clear()
  }

  private queueWorkflowRun(input: QueuedWorkbookAgentWorkflowRun): void {
    const existingTask = this.workflowRunTasks.get(input.sessionState.threadId) ?? Promise.resolve()
    const nextTask = (async () => {
      try {
        await existingTask
      } catch {
        // Continue draining the queue after a prior workflow failure.
      }
      await this.executeQueuedWorkflowRun(input)
    })()
    this.workflowRunTasks.set(input.sessionState.threadId, nextTask)
    void (async () => {
      try {
        await nextTask
      } catch (error) {
        console.error(error)
      } finally {
        if (this.workflowRunTasks.get(input.sessionState.threadId) === nextTask) {
          this.workflowRunTasks.delete(input.sessionState.threadId)
        }
      }
    })()
  }

  private async executeQueuedWorkflowRun(input: QueuedWorkbookAgentWorkflowRun): Promise<void> {
    const currentRun = input.sessionState.durable.workflowRuns.find((run) => run.runId === input.runId)
    if (!currentRun || currentRun.status !== 'running') {
      return
    }
    const abortController = new AbortController()
    this.workflowAbortControllers.set(input.runId, abortController)

    try {
      const result = await executeWorkbookAgentWorkflow({
        documentId: input.documentId,
        zeroSyncService: this.options.zeroSyncService,
        workflowTemplate: input.workflowTemplate,
        context: input.sessionState.durable.context,
        workflowInput: input.workflowInput,
        signal: abortController.signal,
      })
      throwIfWorkflowCancelled(abortController.signal)

      const latestRun = input.sessionState.durable.workflowRuns.find((run) => run.runId === input.runId)
      if (latestRun?.status === 'cancelled') {
        return
      }

      const completedAtUnixMs = this.options.now()
      let completedSummary = result.summary
      const completedRunBase: WorkbookAgentWorkflowRun = {
        ...input.runningRun,
        title: result.title,
        summary: completedSummary,
        status: 'completed',
        updatedAtUnixMs: completedAtUnixMs,
        completedAtUnixMs,
        artifact: result.artifact,
        steps: completeWorkflowSteps(input.workflowTemplate, result.steps, completedAtUnixMs, input.workflowInput),
      }
      if (result.commands && result.commands.length > 0) {
        const baseRevision = await this.options.zeroSyncService.getWorkbookHeadRevision(input.documentId)
        const workflowBundle = attachSharedReviewState(
          createWorkbookAgentCommandBundle({
            documentId: input.documentId,
            threadId: input.sessionState.threadId,
            turnId: input.workflowTurnId,
            goalText: result.goalText ?? result.title,
            baseRevision,
            context: toContextRef(input.sessionState.durable.context),
            commands: result.commands,
            now: completedAtUnixMs,
          }),
          input.sessionState,
        )
        const shouldApplyImmediately = this.options.shouldApplyBundleImmediately(input.sessionState, workflowBundle)
        if (shouldApplyImmediately) {
          const executionRecord = await this.options.applyCommandBundleAutomatically({
            sessionState: input.sessionState,
            actorUserId: input.startedByUserId,
            bundle: workflowBundle,
          })
          if (executionRecord) {
            completedSummary = `Applied workflow: ${executionRecord.summary}`
          }
        } else {
          if (input.sessionState.scope === 'private') {
            throw new Error(
              'Private workbook threads execute workflow changes directly and do not queue review items under the current execution policy.',
            )
          }
          this.options.stageReviewBundle(input.sessionState, input.workflowTurnId, workflowBundle)
        }
      }
      const completedRun: WorkbookAgentWorkflowRun = {
        ...completedRunBase,
        summary: completedSummary,
      }
      input.sessionState.durable.workflowRuns = upsertWorkflowRun(input.sessionState.durable.workflowRuns, completedRun)
      input.sessionState.durable.entries = upsertEntry(
        input.sessionState.durable.entries,
        createSystemEntry(
          `system-workflow-complete:${input.runId}`,
          input.workflowTurnId,
          `Completed workflow: ${result.title}`,
          result.citations,
        ),
      )
      this.options.touch(input.sessionState)
      this.options.incrementCounter('workflowCompletedCount')
      await this.options.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, completedRun)
      await this.options.persistSessionState(input.sessionState)
      this.options.emitSnapshot(input.sessionState.threadId)
      return
    } catch (error) {
      if (isWorkflowAbortError(error)) {
        return
      }
      const latestRun = input.sessionState.durable.workflowRuns.find((run) => run.runId === input.runId)
      if (latestRun?.status === 'cancelled') {
        return
      }
      const failedAtUnixMs = this.options.now()
      const errorMessage = error instanceof Error ? error.message : String(error)
      const failedRun: WorkbookAgentWorkflowRun = {
        ...input.runningRun,
        status: 'failed',
        summary: `Workflow failed: ${input.runningRun.title}`,
        updatedAtUnixMs: failedAtUnixMs,
        completedAtUnixMs: failedAtUnixMs,
        errorMessage,
        steps: failWorkflowSteps(input.workflowTemplate, input.runningRun.steps, errorMessage, failedAtUnixMs, input.workflowInput),
        artifact: null,
      }
      input.sessionState.durable.workflowRuns = upsertWorkflowRun(input.sessionState.durable.workflowRuns, failedRun)
      input.sessionState.durable.entries = upsertEntry(
        input.sessionState.durable.entries,
        createSystemEntry(
          `system-workflow-failed:${input.runId}`,
          input.workflowTurnId,
          failedRun.errorMessage ?? `Workflow failed: ${input.runningRun.title}`,
        ),
      )
      this.options.touch(input.sessionState)
      this.options.incrementCounter('workflowFailedCount')
      await this.options.zeroSyncService.upsertWorkbookWorkflowRun(input.documentId, failedRun)
      await this.options.persistSessionState(input.sessionState)
      this.options.emitSnapshot(input.sessionState.threadId)
      return
    } finally {
      this.workflowAbortControllers.delete(input.runId)
    }
  }
}
