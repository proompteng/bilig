import type { WorkbookAgentAppliedBy } from '@bilig/agent-api'
import type { WorkbookAgentStreamEvent, WorkbookAgentThreadSnapshot, WorkbookAgentThreadSummary } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import {
  createDisabledWorkbookAgentObservabilitySnapshot,
  type WorkbookAgentObservabilitySnapshot,
} from './workbook-agent-session-registry.js'
import type { WorkbookAgentService } from './workbook-agent-service-options.js'

export class DisabledWorkbookAgentService implements WorkbookAgentService {
  readonly enabled = false

  async createSession(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async updateContext(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async startTurn(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async startWorkflow(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async cancelWorkflow(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async interruptTurn(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async applyReviewItem(_input: {
    documentId: string
    threadId: string
    reviewItemId: string
    session: SessionIdentity
    appliedBy: WorkbookAgentAppliedBy
    commandIndexes?: readonly number[] | null
    preview: unknown
  }): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async reviewReviewItem(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async dismissReviewItem(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async replayExecutionRecord(): Promise<never> {
    throw new Error('Workbook agent service is not configured')
  }

  async listThreads(_input: { documentId: string; session: SessionIdentity }): Promise<WorkbookAgentThreadSummary[]> {
    throw new Error('Workbook agent service is not configured')
  }

  getObservabilitySnapshot(): WorkbookAgentObservabilitySnapshot {
    return createDisabledWorkbookAgentObservabilitySnapshot(Date.now())
  }

  getSnapshot(_input: { documentId: string; threadId: string; session: SessionIdentity }): WorkbookAgentThreadSnapshot {
    throw new Error('Workbook agent service is not configured')
  }

  subscribe(_threadId: string, _listener: (event: WorkbookAgentStreamEvent) => void): () => void {
    return () => {}
  }

  async close(): Promise<void> {}
}
