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

function persistenceKey(input: {
  documentId: string;
  threadId: string;
  actorUserId: string;
}): string {
  return `${input.documentId}\u0000${input.threadId}\u0000${input.actorUserId}`;
}

function dedupeTimelineEntries(
  entries: readonly WorkbookAgentTimelineEntry[],
): WorkbookAgentTimelineEntry[] {
  const deduped: WorkbookAgentTimelineEntry[] = [];
  const indexById = new Map<string, number>();
  for (const entry of entries) {
    const existingIndex = indexById.get(entry.id);
    if (existingIndex === undefined) {
      indexById.set(entry.id, deduped.length);
      deduped.push(entry);
      continue;
    }
    deduped[existingIndex] = entry;
  }
  return deduped;
}

export class WorkbookAgentSessionStore {
  private readonly pendingSaves = new Map<string, Promise<void>>();

  constructor(private readonly source: WorkbookAgentThreadPersistenceSource) {}

  async saveSessionSnapshot(input: WorkbookAgentPersistedSessionInput): Promise<void> {
    const key = persistenceKey(input);
    const entries = dedupeTimelineEntries(input.entries);
    const previous = this.pendingSaves.get(key) ?? Promise.resolve();
    const next = previous.catch(() => undefined).then(async () => {
      await this.source.saveWorkbookAgentThreadState({
        documentId: input.documentId,
        threadId: input.threadId,
        actorUserId: input.actorUserId,
        scope: input.scope,
        context: input.context,
        entries,
        pendingBundle: input.pendingBundle,
        updatedAtUnixMs: input.updatedAtUnixMs,
      });
      return undefined;
    });
    this.pendingSaves.set(key, next);
    try {
      await next;
    } finally {
      if (this.pendingSaves.get(key) === next) {
        this.pendingSaves.delete(key);
      }
    }
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
