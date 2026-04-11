import {
  appendWorkbookAgentCommandToBundle,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type WorkbookAgentExecutionRecord,
} from "@bilig/agent-api";
import type { WorkbookAgentSessionSnapshot } from "@bilig/contracts";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  handleWorkbookAgentToolCall,
  type WorkbookAgentStartWorkflowRequest,
} from "./workbook-agent-tools.js";
import { type WorkbookAgentSessionState, toContextRef } from "./workbook-agent-service-shared.js";

export function createWorkbookAgentDynamicToolHandler(input: {
  zeroSyncService: ZeroSyncService;
  now: () => number;
  getSessionByThreadId: (threadId: string) => WorkbookAgentSessionState;
  resolveTurnActorUserId: (sessionState: WorkbookAgentSessionState, turnId: string) => string;
  resolveTurnContext: (
    sessionState: WorkbookAgentSessionState,
    turnId: string,
  ) => WorkbookAgentSessionState["snapshot"]["context"];
  stagePendingBundle: (
    sessionState: WorkbookAgentSessionState,
    turnId: string,
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>,
  ) => void;
  shouldApplyToolCommandImmediately: (
    sessionState: WorkbookAgentSessionState,
    command: ReturnType<typeof appendWorkbookAgentCommandToBundle>["commands"][number],
  ) => boolean;
  applyToolCommandImmediately: (input: {
    sessionState: WorkbookAgentSessionState;
    actorUserId: string;
    bundle: ReturnType<typeof appendWorkbookAgentCommandToBundle>;
  }) => Promise<WorkbookAgentExecutionRecord | null>;
  persistSessionState: (sessionState: WorkbookAgentSessionState) => Promise<void>;
  emitSnapshot: (threadId: string) => void;
  startWorkflow: (input: {
    documentId: string;
    sessionId: string;
    session: SessionIdentity;
    body: WorkbookAgentStartWorkflowRequest & {
      context?: WorkbookAgentSessionState["snapshot"]["context"];
    };
  }) => Promise<WorkbookAgentSessionSnapshot>;
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
          const baseRevision = await input.zeroSyncService.getWorkbookHeadRevision(
            sessionState.documentId,
          );
          const bundle = appendWorkbookAgentCommandToBundle({
            previousBundle: sessionState.snapshot.pendingBundle,
            documentId: sessionState.documentId,
            threadId: sessionState.threadId,
            turnId: request.turnId,
            goalText:
              sessionState.promptByTurn.get(request.turnId) ??
              "Update workbook from assistant request",
            baseRevision,
            context: toContextRef(requestContext),
            command,
            now: input.now(),
          });
          if (input.shouldApplyToolCommandImmediately(sessionState, command)) {
            try {
              const executionRecord = await input.applyToolCommandImmediately({
                sessionState,
                actorUserId: requestActorUserId,
                bundle,
              });
              return {
                bundle,
                executionRecord,
              };
            } catch {
              input.stagePendingBundle(sessionState, request.turnId, bundle);
              await input.persistSessionState(sessionState);
              input.emitSnapshot(sessionState.threadId);
              return bundle;
            }
          }
          input.stagePendingBundle(sessionState, request.turnId, bundle);
          await input.persistSessionState(sessionState);
          input.emitSnapshot(sessionState.threadId);
          return bundle;
        },
        startWorkflow: async (workflowRequest: WorkbookAgentStartWorkflowRequest) => {
          const previousRunIds = new Set(
            sessionState.snapshot.workflowRuns.map((run) => run.runId),
          );
          const nextSnapshot = await input.startWorkflow({
            documentId: sessionState.documentId,
            sessionId: sessionState.sessionId,
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
