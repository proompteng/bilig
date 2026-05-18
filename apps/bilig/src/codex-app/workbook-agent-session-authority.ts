import { createWorkbookAgentServiceError } from '../workbook-agent-errors.js'
import type { WorkbookAgentFeatureFlags } from './workbook-agent-feature-flags.js'
import type { WorkbookAgentSessionRegistry } from './workbook-agent-session-registry.js'
import type { WorkbookAgentThreadRepository } from './workbook-agent-thread-repository.js'
import { assertWorkbookAgentSessionAccessPolicy } from './workbook-agent-service-access-policy.js'
import type { WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export class WorkbookAgentSessionAuthority {
  constructor(
    private readonly input: {
      readonly featureFlags: () => WorkbookAgentFeatureFlags
      readonly sessionRegistry: WorkbookAgentSessionRegistry
      readonly threadRepository: WorkbookAgentThreadRepository
    },
  ) {}

  getOwnedSession(documentId: string, threadId: string, userId: string): WorkbookAgentThreadState {
    const sessionState = this.input.sessionRegistry.tryGetSession(threadId)
    if (!sessionState) {
      throw createWorkbookAgentServiceError({
        code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
        message: 'Workbook agent thread not found',
        statusCode: 404,
        retryable: true,
      })
    }
    return this.requireOwnedSession(sessionState, documentId, userId)
  }

  async getAuthorizedSession(documentId: string, threadId: string, userId: string): Promise<WorkbookAgentThreadState> {
    const sessionState = this.getOwnedSession(documentId, threadId, userId)
    await this.authorizeSharedSessionForUser(sessionState, documentId, userId)
    return sessionState
  }

  requireOwnedSession(sessionState: WorkbookAgentThreadState, documentId: string, userId: string): WorkbookAgentThreadState {
    if (sessionState.documentId !== documentId) {
      throw this.createHiddenThreadError()
    }
    if (sessionState.scope !== 'shared' && sessionState.userId !== userId) {
      throw this.createHiddenThreadError()
    }
    assertWorkbookAgentSessionAccessPolicy({
      featureFlags: this.input.featureFlags(),
      sessionState,
      documentId,
      userId,
    })
    return sessionState
  }

  assertSharedSessionAlreadyAuthorized(sessionState: WorkbookAgentThreadState, userId: string): void {
    if (sessionState.scope !== 'shared' || sessionState.live.authorizedUserIds.has(userId)) {
      return
    }
    throw this.createHiddenThreadError()
  }

  async authorizeSharedSessionForUser(sessionState: WorkbookAgentThreadState, documentId: string, userId: string): Promise<void> {
    if (sessionState.scope !== 'shared' || sessionState.live.authorizedUserIds.has(userId)) {
      return
    }
    const durableThreadSession = await this.input.threadRepository.loadThreadState({
      documentId,
      actorUserId: userId,
      threadId: sessionState.threadId,
    })
    if (durableThreadSession.threadState?.scope !== 'shared') {
      throw this.createHiddenThreadError()
    }
    sessionState.live.authorizedUserIds.add(userId)
  }

  getSessionByThreadId(threadId: string): WorkbookAgentThreadState {
    const sessionState = this.tryGetSessionByThreadId(threadId)
    if (!sessionState) {
      throw new Error(`Workbook agent thread not found for thread ${threadId}`)
    }
    return sessionState
  }

  tryGetSessionByThreadId(threadId: string): WorkbookAgentThreadState | null {
    return this.input.sessionRegistry.tryGetSession(threadId)
  }

  private createHiddenThreadError(): Error {
    return createWorkbookAgentServiceError({
      code: 'WORKBOOK_AGENT_THREAD_NOT_FOUND',
      message: 'Workbook agent thread not found',
      statusCode: 404,
      retryable: false,
    })
  }
}
