import { describe, expect, it, vi } from "vitest";
import type { WorkbookAgentExecutionRecord } from "@bilig/agent-api";
import type { WorkbookAgentWorkflowRun } from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import { createSystemEntry } from "./workbook-agent-session-model.js";
import { WorkbookAgentThreadRepository } from "./workbook-agent-thread-repository.js";

function createPersistenceSource(
  overrides: Partial<
    Pick<
      ZeroSyncService,
      | "loadWorkbookAgentThreadState"
      | "saveWorkbookAgentThreadState"
      | "listWorkbookAgentThreadRuns"
      | "listWorkbookThreadWorkflowRuns"
    >
  > = {},
): Pick<
  ZeroSyncService,
  | "loadWorkbookAgentThreadState"
  | "saveWorkbookAgentThreadState"
  | "listWorkbookAgentThreadRuns"
  | "listWorkbookThreadWorkflowRuns"
> {
  return {
    async loadWorkbookAgentThreadState() {
      return null;
    },
    async saveWorkbookAgentThreadState() {},
    async listWorkbookAgentThreadRuns() {
      return [];
    },
    async listWorkbookThreadWorkflowRuns() {
      return [];
    },
    ...overrides,
  };
}

describe("WorkbookAgentThreadRepository", () => {
  it("persists the canonical thread snapshot through ZeroSync", async () => {
    const saveWorkbookAgentThreadState = vi.fn(async () => {});
    const store = new WorkbookAgentThreadRepository(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    await store.saveThreadState({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      executionPolicy: "ownerReview",
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "B2",
          range: {
            startAddress: "B2",
            endAddress: "D5",
          },
        },
        viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
      },
      entries: [createSystemEntry("entry-1", null, "hello")],
      reviewQueueItems: [],
      updatedAtUnixMs: 123,
    });

    expect(saveWorkbookAgentThreadState).toHaveBeenCalledWith({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      executionPolicy: "ownerReview",
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "B2",
          range: {
            startAddress: "B2",
            endAddress: "D5",
          },
        },
        viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
      },
      entries: [createSystemEntry("entry-1", null, "hello")],
      reviewQueueItems: [],
      updatedAtUnixMs: 123,
    });
  });

  it("serializes overlapping saves for the same durable thread", async () => {
    let resolveFirstSave!: () => void;
    const firstSaveBlocked = new Promise<void>((resolve) => {
      resolveFirstSave = resolve;
    });
    const saveOrder: number[] = [];
    const saveWorkbookAgentThreadState = vi.fn(async (record) => {
      saveOrder.push(record.updatedAtUnixMs);
      if (record.updatedAtUnixMs !== 100) {
        return;
      }
      await firstSaveBlocked;
    });
    const store = new WorkbookAgentThreadRepository(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    const firstSave = store.saveThreadState({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      executionPolicy: "ownerReview",
      context: null,
      entries: [createSystemEntry("entry-1", null, "first")],
      reviewQueueItems: [],
      updatedAtUnixMs: 100,
    });
    const secondSave = store.saveThreadState({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      executionPolicy: "ownerReview",
      context: null,
      entries: [createSystemEntry("entry-2", null, "second")],
      reviewQueueItems: [],
      updatedAtUnixMs: 200,
    });

    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(saveWorkbookAgentThreadState).toHaveBeenCalledTimes(1);

    resolveFirstSave();
    await Promise.all([firstSave, secondSave]);

    expect(saveOrder).toEqual([100, 200]);
  });

  it("dedupes duplicate entry ids before persisting", async () => {
    const saveWorkbookAgentThreadState = vi.fn(async () => {});
    const store = new WorkbookAgentThreadRepository(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    await store.saveThreadState({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      executionPolicy: "ownerReview",
      context: null,
      entries: [
        createSystemEntry("entry-1", null, "first"),
        createSystemEntry("entry-1", null, "second"),
      ],
      reviewQueueItems: [],
      updatedAtUnixMs: 123,
    });

    expect(saveWorkbookAgentThreadState).toHaveBeenCalledWith(
      expect.objectContaining({
        entries: [createSystemEntry("entry-1", null, "second")],
      }),
    );
  });

  it("loads durable thread state, execution records, and workflow runs together", async () => {
    const executionRecord: WorkbookAgentExecutionRecord = {
      id: "run-1",
      bundleId: "bundle-1",
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      actorUserId: "alex@example.com",
      goalText: "Normalize imported rows",
      planText: null,
      summary: "Write cells in Sheet1!B2",
      scope: "sheet",
      riskClass: "medium",
      approvalMode: "preview",
      acceptedScope: "full",
      appliedBy: "user",
      baseRevision: 3,
      appliedRevision: 4,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 110,
      context: null,
      commands: [
        {
          kind: "writeRange",
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      ],
      preview: null,
    };
    const workflowRun: WorkbookAgentWorkflowRun = {
      runId: "workflow-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeWorkbook",
      title: "Summarize Workbook",
      summary: "Summarized the workbook",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 140,
      completedAtUnixMs: 140,
      errorMessage: null,
      steps: [],
      artifact: null,
    };
    const loadWorkbookAgentThreadState = vi.fn(async () => ({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "private" as const,
      executionPolicy: "autoApplyAll" as const,
      context: null,
      entries: [],
      reviewQueueItems: [],
      updatedAtUnixMs: 100,
    }));
    const listWorkbookAgentThreadRuns = vi.fn(async () => [executionRecord]);
    const listWorkbookThreadWorkflowRuns = vi.fn(async () => [workflowRun]);
    const store = new WorkbookAgentThreadRepository(
      createPersistenceSource({
        loadWorkbookAgentThreadState,
        listWorkbookAgentThreadRuns,
        listWorkbookThreadWorkflowRuns,
      }),
    );

    const loaded = await store.loadThreadState({
      documentId: "doc-1",
      actorUserId: "alex@example.com",
      threadId: "thr-1",
    });

    expect(loadWorkbookAgentThreadState).toHaveBeenCalledWith("doc-1", "alex@example.com", "thr-1");
    expect(listWorkbookAgentThreadRuns).toHaveBeenCalledWith("doc-1", "alex@example.com", "thr-1");
    expect(listWorkbookThreadWorkflowRuns).toHaveBeenCalledWith(
      "doc-1",
      "alex@example.com",
      "thr-1",
    );
    expect(loaded).toEqual({
      threadState: {
        documentId: "doc-1",
        threadId: "thr-1",
        actorUserId: "alex@example.com",
        scope: "private",
        executionPolicy: "autoApplyAll",
        context: null,
        entries: [],
        reviewQueueItems: [],
        updatedAtUnixMs: 100,
      },
      executionRecords: [executionRecord],
      workflowRuns: [workflowRun],
    });
  });
});
