import type { WorkbookAgentCommandBundle, WorkbookAgentExecutionRecord } from "@bilig/agent-api";
import type {
  WorkbookAgentTimelineEntry,
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
} from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookAgentThreadStateRecord } from "../zero/workbook-chat-thread-store.js";

type WorkbookAgentThreadPersistenceSource = Pick<
  ZeroSyncService,
  | "loadWorkbookAgentThreadState"
  | "saveWorkbookAgentThreadState"
  | "listWorkbookAgentThreadRuns"
  | "listWorkbookThreadWorkflowRuns"
>;

export interface WorkbookAgentPersistedSessionInput {
  readonly documentId: string;
  readonly threadId: string;
  readonly actorUserId: string;
  readonly scope: "private" | "shared";
  readonly context: WorkbookAgentUiContext | null;
  readonly entries: readonly WorkbookAgentTimelineEntry[];
  readonly pendingBundle: WorkbookAgentCommandBundle | null;
  readonly updatedAtUnixMs: number;
}

export interface WorkbookAgentLoadedThreadSession {
  readonly threadState: WorkbookAgentThreadStateRecord | null;
  readonly executionRecords: WorkbookAgentExecutionRecord[];
  readonly workflowRuns: WorkbookAgentWorkflowRun[];
}

export class WorkbookAgentSessionStore {
  constructor(private readonly source: WorkbookAgentThreadPersistenceSource) {}

  async saveSessionSnapshot(input: WorkbookAgentPersistedSessionInput): Promise<void> {
    await this.source.saveWorkbookAgentThreadState({
      documentId: input.documentId,
      threadId: input.threadId,
      actorUserId: input.actorUserId,
      scope: input.scope,
      context: input.context,
      entries: [...input.entries],
      pendingBundle: input.pendingBundle,
      updatedAtUnixMs: input.updatedAtUnixMs,
    });
  }

  async loadThreadSession(input: {
    documentId: string;
    actorUserId: string;
    threadId: string;
  }): Promise<WorkbookAgentLoadedThreadSession> {
    const [threadState, executionRecords, workflowRuns] = await Promise.all([
      this.source.loadWorkbookAgentThreadState(input.documentId, input.actorUserId, input.threadId),
      this.source.listWorkbookAgentThreadRuns(input.documentId, input.actorUserId, input.threadId),
      this.source.listWorkbookThreadWorkflowRuns(
        input.documentId,
        input.actorUserId,
        input.threadId,
      ),
    ]);
    return {
      threadState,
      executionRecords,
      workflowRuns,
    };
  }
}
