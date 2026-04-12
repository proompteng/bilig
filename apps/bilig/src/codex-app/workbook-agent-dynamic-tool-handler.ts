import {
  appendWorkbookAgentCommandToBundle,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type WorkbookAgentExecutionRecord,
} from "@bilig/agent-api";
import type { WorkbookAgentThreadSnapshot } from "@bilig/contracts";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  handleWorkbookAgentToolCall,
  type WorkbookAgentStartWorkflowRequest,
} from "./workbook-agent-tools.js";
import { type WorkbookAgentThreadState, toContextRef } from "./workbook-agent-service-shared.js";

export function createWorkbookAgentDynamicToolHandler(input: {
  zeroSyncService: ZeroSyncService;
  now: () => number;
  getSessionByThreadId: (threadId: string) => WorkbookAgentThreadState;
  resolveTurnActorUserId: (sessionState: WorkbookAgentThreadState, turnId: string) => string;
  resolveTurnContext: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
  ) => WorkbookAgentThreadState["durable"]["context"];
  stageReviewBundle: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => void;
  queuePrivateTurnBundle: (
    sessionState: WorkbookAgentThreadState,
    turnId: string,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => void;
  shouldApplyToolBundleImmediately: (
    sessionState: WorkbookAgentThreadState,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => boolean;
  applyToolBundleAutomatically: (input: {
    sessionState: WorkbookAgentThreadState;
    actorUserId: string;
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>;
  }) => Promise<WorkbookAgentExecutionRecord | null>;
  persistSessionState: (sessionState: WorkbookAgentThreadState) => Promise<void>;
  emitSnapshot: (threadId: string) => void;
  startWorkflow: (input: {
    documentId: string;
    threadId: string;
    session: SessionIdentity;
    body: WorkbookAgentStartWorkflowRequest & {
      context?: WorkbookAgentThreadState["durable"]["context"];
    };
  }) => Promise<WorkbookAgentThreadSnapshot>;
}): (request: CodexDynamicToolCallRequest) => Promise<CodexDynamicToolCallResult> {
  return async (request: CodexDynamicToolCallRequest) => {
    const sessionState = input.getSessionByThreadId(request.threadId);
    const requestActorUserId = input.resolveTurnActorUserId(sessionState, request.turnId);
    const requestContext = input.resolveTurnContext(sessionState, request.turnId);
    return handleWorkbookAgentToolCall(
      {
        documentId: sessionState.documentId,
        session: {
          userID: requestActorUserId,
          roles: ["editor"],
        },
        uiContext: requestContext,
        zeroSyncService: input.zeroSyncService,
        stageCommand: async (command) => {
          const previousBundle =
            sessionState.scope === "private"
              ? (sessionState.live.stagedPrivateBundleByTurn.get(request.turnId) ??
                sessionState.durable.pendingBundle)
              : sessionState.durable.pendingBundle;
          const baseRevision = await input.zeroSyncService.getWorkbookHeadRevision(
            sessionState.documentId,
          );
          const bundle = appendWorkbookAgentCommandToBundle({
            previousBundle,
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            turnId: request.turnId,
            goalText:
              sessionState.live.promptByTurn.get(request.turnId) ??
              "Update workbook from assistant request",
            baseRevision,
            context: toContextRef(requestContext),
            command,
            now: input.now(),
          });
          if (
            sessionState.scope === "private" &&
            input.shouldApplyToolBundleImmediately(sessionState, bundle)
          ) {
            input.queuePrivateTurnBundle(sessionState, request.turnId, bundle);
            return {
              bundle,
              executionRecord: null,
              disposition: "queuedForTurnApply",
            };
          }
          if (input.shouldApplyToolBundleImmediately(sessionState, bundle)) {
            const executionRecord = await input.applyToolBundleAutomatically({
              sessionState,
              actorUserId: requestActorUserId,
              bundle,
            });
            return {
              bundle,
              executionRecord,
            };
          }
          input.stageReviewBundle(sessionState, request.turnId, bundle);
          await input.persistSessionState(sessionState);
          input.emitSnapshot(sessionState.threadId);
          return {
            bundle,
            executionRecord: null,
            disposition: "reviewQueued",
          };
        },
        startWorkflow: async (workflowRequest: WorkbookAgentStartWorkflowRequest) => {
          const previousRunIds = new Set(sessionState.durable.workflowRuns.map((run) => run.runId));
          const nextSnapshot = await input.startWorkflow({
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            session: {
              userID: requestActorUserId,
              roles: ["editor"],
            },
            body: {
              ...workflowRequest,
              ...(requestContext ? { context: requestContext } : {}),
            },
          });
          const nextRun =
            nextSnapshot.workflowRuns.find((run) => !previousRunIds.has(run.runId)) ??
            nextSnapshot.workflowRuns.find(
              (run) => run.workflowTemplate === workflowRequest.workflowTemplate,
            ) ??
            null;
          if (!nextRun) {
            throw new Error(
              `Workflow run not found after starting ${workflowRequest.workflowTemplate}`,
            );
          }
          return nextRun;
        },
      },
      request,
    );
  };
}
