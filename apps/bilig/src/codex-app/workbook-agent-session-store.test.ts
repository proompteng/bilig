import { describe, expect, it, vi } from "vitest";
import type { WorkbookAgentExecutionRecord } from "@bilig/agent-api";
import type { WorkbookAgentWorkflowRun } from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import { createSystemEntry } from "./workbook-agent-session-model.js";
import { WorkbookAgentSessionStore } from "./workbook-agent-session-store.js";

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

describe("WorkbookAgentSessionStore", () => {
  it("persists the canonical thread snapshot through ZeroSync", async () => {
    const saveWorkbookAgentThreadState = vi.fn(async () => {});
    const store = new WorkbookAgentSessionStore(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    await store.saveSessionSnapshot({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: {
        selection: { sheetName: "Sheet1", address: "B2" },
        viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
      },
      entries: [createSystemEntry("entry-1", null, "hello")],
      pendingBundle: null,
      updatedAtUnixMs: 123,
    });

    expect(saveWorkbookAgentThreadState).toHaveBeenCalledWith({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: {
        selection: { sheetName: "Sheet1", address: "B2" },
        viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
      },
      entries: [createSystemEntry("entry-1", null, "hello")],
      pendingBundle: null,
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
    const store = new WorkbookAgentSessionStore(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    const firstSave = store.saveSessionSnapshot({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: null,
      entries: [createSystemEntry("entry-1", null, "first")],
      pendingBundle: null,
      updatedAtUnixMs: 100,
    });
    const secondSave = store.saveSessionSnapshot({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: null,
      entries: [createSystemEntry("entry-2", null, "second")],
      pendingBundle: null,
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
    const store = new WorkbookAgentSessionStore(
      createPersistenceSource({
        saveWorkbookAgentThreadState,
      }),
    );

    await store.saveSessionSnapshot({
      documentId: "doc-1",
      threadId: "thr-1",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: null,
      entries: [
        createSystemEntry("entry-1", null, "first"),
        createSystemEntry("entry-1", null, "second"),
      ],
      pendingBundle: null,
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
      context: null,
      entries: [],
      pendingBundle: null,
      updatedAtUnixMs: 100,
    }));
    const listWorkbookAgentThreadRuns = vi.fn(async () => [executionRecord]);
    const listWorkbookThreadWorkflowRuns = vi.fn(async () => [workflowRun]);
    const store = new WorkbookAgentSessionStore(
      createPersistenceSource({
        loadWorkbookAgentThreadState,
        listWorkbookAgentThreadRuns,
        listWorkbookThreadWorkflowRuns,
      }),
    );

    const loaded = await store.loadThreadSession({
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
        context: null,
        entries: [],
        pendingBundle: null,
        updatedAtUnixMs: 100,
      },
      executionRecords: [executionRecord],
      workflowRuns: [workflowRun],
    });
  });
});
