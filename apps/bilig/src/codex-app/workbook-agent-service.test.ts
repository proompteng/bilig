import {
  isWorkbookAgentCommandBundle,
  isWorkbookAgentExecutionRecord,
  type CodexServerNotification,
  type CodexTurn,
} from "@bilig/agent-api";
import { SpreadsheetEngine } from "@bilig/core";
import { describe, expect, it, vi } from "vitest";
import type { ZeroSyncService } from "../zero/service.js";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";
import type { WorkbookAgentThreadStateRecord } from "../zero/workbook-chat-thread-store.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import type {
  CodexAppServerClientOptions,
  CodexAppServerTransport,
} from "./codex-app-server-client.js";
import { createWorkbookAgentService } from "./workbook-agent-service.js";
import { createWorkbookAgentServiceError } from "../workbook-agent-errors.js";

class FakeCodexTransport implements CodexAppServerTransport {
  private readonly listeners = new Set<(notification: CodexServerNotification) => void>();
  private turnCounter = 0;
  lastThreadStartInput: Parameters<CodexAppServerTransport["threadStart"]>[0] | null = null;
  lastThreadResumeInput: { threadId: string } | null = null;
  resumeError: unknown = null;

  async ensureReady() {
    return {
      userAgent: "fake",
      codexHome: "/tmp/fake-codex",
      platformFamily: "unix",
      platformOs: "macos",
    };
  }

  subscribe(listener: (notification: CodexServerNotification) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  async threadStart(input: Parameters<CodexAppServerTransport["threadStart"]>[0]) {
    this.lastThreadStartInput = input;
    return {
      id: "thr-test",
      preview: "",
      turns: [],
    };
  }

  async threadResume(input: { threadId: string }) {
    this.lastThreadResumeInput = input;
    if (this.resumeError) {
      throw this.resumeError;
    }
    return {
      id: input.threadId,
      preview: "",
      turns: [],
    };
  }

  async turnStart(): Promise<CodexTurn> {
    this.turnCounter += 1;
    return {
      id: `turn-${String(this.turnCounter)}`,
      status: "inProgress",
      items: [],
      error: null,
    };
  }

  async turnInterrupt() {}

  async close() {}

  emit(notification: CodexServerNotification): void {
    this.listeners.forEach((listener) => listener(notification));
  }
}

function createPreviewSummary(overrides: Record<string, unknown> = {}) {
  return {
    ranges: [],
    structuralChanges: [],
    cellDiffs: [],
    effectSummary: {
      displayedCellDiffCount: 0,
      truncatedCellDiffs: false,
      inputChangeCount: 0,
      formulaChangeCount: 0,
      styleChangeCount: 0,
      numberFormatChangeCount: 0,
      structuralChangeCount: 0,
    },
    ...overrides,
  };
}

function createZeroSyncStub(overrides: Partial<ZeroSyncService> = {}): ZeroSyncService {
  return {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error("not used");
    },
    async handleMutate() {
      throw new Error("not used");
    },
    async inspectWorkbook<T>(_documentId: string, _task: (runtime: never) => T | Promise<T>) {
      throw new Error("not used");
    },
    async applyServerMutator() {},
    async applyAgentCommandBundle() {
      return { revision: 2, preview: createPreviewSummary() };
    },
    async listWorkbookChanges() {
      return [];
    },
    async listWorkbookAgentRuns() {
      return [];
    },
    async listWorkbookAgentThreadRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {},
    async listWorkbookAgentThreadSummaries() {
      return [];
    },
    async loadWorkbookAgentThreadState() {
      return null;
    },
    async saveWorkbookAgentThreadState() {},
    async listWorkbookThreadWorkflowRuns() {
      return [];
    },
    async upsertWorkbookWorkflowRun() {},
    async getWorkbookHeadRevision() {
      return 1;
    },
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
    ...overrides,
  };
}

describe("workbook agent service", () => {
  it("boots the Codex app-server transport with local workbook skills", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: {
      current: CodexAppServerClientOptions | null;
    } = { current: null };
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
        capturedOptions.current = options;
        return fakeCodex;
      },
    });

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      expect(capturedOptions.current?.args).toEqual([
        "app-server",
        "-c",
        "analytics.enabled=false",
      ]);
      expect(fakeCodex.lastThreadStartInput?.dynamicTools.map((tool) => tool.name)).toEqual(
        expect.arrayContaining([
          "bilig_read_selection",
          "bilig_read_visible_range",
          "bilig_start_workflow",
          "bilig_inspect_cell",
          "bilig_find_formula_issues",
          "bilig_search_workbook",
          "bilig_trace_dependencies",
          "bilig_read_range",
          "bilig_write_range",
        ]),
      );
      expect(
        fakeCodex.lastThreadStartInput?.dynamicTools.every((tool) =>
          /^[a-zA-Z0-9_-]+$/.test(tool.name),
        ),
      ).toBe(true);
      expect(fakeCodex.lastThreadStartInput?.baseInstructions).toContain("bilig workbook tools");
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        "bilig_read_workbook",
      );
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        "bilig_start_workflow",
      );
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        "bilig_search_workbook",
      );
    } finally {
      await service.close();
    }
  });

  it("streams assistant updates into the session timeline", async () => {
    const fakeCodex = new FakeCodexTransport();
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      expect(snapshot.threadId).toBe("thr-test");
      expect(snapshot.pendingBundle).toBeNull();
      expect(snapshot.executionRecords).toEqual([]);

      const events: unknown[] = [];
      const unsubscribe = service.subscribe(snapshot.threadId, (event) => {
        events.push(event);
      });

      const inProgress = await service.startTurn({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          prompt: "Summarize Sheet1",
        },
      });

      expect(inProgress.status).toBe("inProgress");
      expect(inProgress.entries.some((entry) => entry.kind === "user")).toBe(true);

      fakeCodex.emit({
        method: "item/agentMessage/delta",
        params: {
          threadId: "thr-test",
          turnId: "turn-1",
          itemId: "msg-1",
          delta: "Checking Sheet1",
        },
      });
      fakeCodex.emit({
        method: "item/completed",
        params: {
          threadId: "thr-test",
          turnId: "turn-1",
          item: {
            type: "agentMessage",
            id: "msg-1",
            text: "Checking Sheet1",
            phase: null,
            memoryCitation: null,
          },
        },
      });
      fakeCodex.emit({
        method: "turn/completed",
        params: {
          threadId: "thr-test",
          turn: {
            id: "turn-1",
            status: "completed",
            items: [],
            error: null,
          },
        },
      });

      const finalSnapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(finalSnapshot.status).toBe("idle");
      expect(finalSnapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            id: "msg-1",
            kind: "assistant",
            text: "Checking Sheet1",
          }),
        ]),
      );
      expect(events).toContainEqual({
        type: "assistantDelta",
        itemId: "msg-1",
        delta: "Checking Sheet1",
      });

      unsubscribe();
    } finally {
      await service.close();
    }
  });

  it("starts durable read/report workflows and records completed runs", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellFormula("Sheet1", "B1", "SUM(A1:A1)");
    let inspectWorkbookCallCount = 0;
    const inspectWorkbook = async <T>(
      _documentId: string,
      task: (runtime: WorkbookRuntime) => T | Promise<T>,
    ): Promise<T> => {
      inspectWorkbookCallCount += 1;
      const runtime: WorkbookRuntime = {
        documentId: "doc-1",
        engine,
        projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: "alex@example.com",
          updatedBy: "alex@example.com",
          updatedAt: "2026-04-10T00:00:00.000Z",
        }),
        headRevision: 1,
        calculatedRevision: 1,
        ownerUserId: "alex@example.com",
      };
      return await task(runtime);
    };
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        inspectWorkbook,
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "summarizeWorkbook",
        },
      });

      expect(inspectWorkbookCallCount).toBe(1);
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "summarizeWorkbook",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "inspect-workbook",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            kind: "markdown",
            title: "Workbook Summary",
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Started workflow: Summarize Workbook",
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Summarize Workbook",
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("runs durable formula issue workflows with cited issue reports", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellFormula("Sheet1", "B1", "1/0");
    engine.setCellFormula("Sheet1", "C1", "LEN(A1:A2)");
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "findFormulaIssues",
          sheetName: "Sheet1",
        },
      });

      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "findFormulaIssues",
          title: "Find Formula Issues",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "scan-formula-cells",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            title: "Formula Issues",
            text: expect.stringContaining("## Formula Issues"),
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Started workflow: Find Formula Issues",
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Find Formula Issues",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Sheet1",
                startAddress: "B1",
                endAddress: "B1",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("stages formula-highlight preview bundles from durable workflows", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellFormula("Sheet1", "B1", "1/0");
    engine.setCellFormula("Sheet1", "C1", "LEN(A1:A2)");
    const getWorkbookHeadRevision = vi.fn(async () => 7);
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        getWorkbookHeadRevision,
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "highlightFormulaIssues",
          sheetName: "Sheet1",
        },
      });

      expect(getWorkbookHeadRevision).toHaveBeenCalledWith("doc-1");
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "highlightFormulaIssues",
          title: "Highlight Formula Issues",
          status: "completed",
          artifact: expect.objectContaining({
            title: "Formula Issue Highlights",
            text: expect.stringContaining("## Highlighted Formula Issues"),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "stage-issue-highlights",
              status: "completed",
            }),
          ]),
        }),
      );
      expect(snapshot.pendingBundle).toEqual(
        expect.objectContaining({
          baseRevision: 7,
          commands: expect.arrayContaining([
            expect.objectContaining({
              kind: "formatRange",
              range: expect.objectContaining({
                sheetName: "Sheet1",
                startAddress: "B1",
                endAddress: "B1",
              }),
              patch: expect.objectContaining({
                fill: expect.objectContaining({
                  backgroundColor: "#FEE2E2",
                }),
              }),
            }),
          ]),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: expect.stringContaining("Staged preview bundle"),
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Highlight Formula Issues",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Sheet1",
                startAddress: "B1",
                endAddress: "B1",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("stages header-normalization preview bundles from durable workflows", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Imports");
    engine.setCellValue("Imports", "A1", "order_id");
    engine.setCellValue("Imports", "B1", " customer name ");
    engine.setCellValue("Imports", "C1", "customer_name");
    engine.setCellValue("Imports", "A2", 1001);
    const getWorkbookHeadRevision = vi.fn(async () => 7);
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        getWorkbookHeadRevision,
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Imports",
              address: "A2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "normalizeCurrentSheetHeaders",
          sheetName: "Imports",
        },
      });

      expect(getWorkbookHeadRevision).toHaveBeenCalledWith("doc-1");
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "normalizeCurrentSheetHeaders",
          title: "Normalize Current Sheet Headers",
          status: "completed",
          artifact: expect.objectContaining({
            title: "Header Normalization Preview",
            text: expect.stringContaining("## Header Normalization Preview"),
          }),
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "stage-header-normalization",
              status: "completed",
            }),
          ]),
        }),
      );
      expect(snapshot.pendingBundle).toEqual(
        expect.objectContaining({
          baseRevision: 7,
          commands: [
            expect.objectContaining({
              kind: "writeRange",
              sheetName: "Imports",
              startAddress: "A1",
              values: [["Order Id", "Customer Name", "Customer Name 2"]],
            }),
          ],
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: expect.stringContaining("Staged preview bundle"),
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Normalize Current Sheet Headers",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Imports",
                startAddress: "A1",
                endAddress: "C1",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("runs durable current-sheet summary workflows from the active selection context", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Revenue");
    engine.setCellValue("Revenue", "A1", "Region");
    engine.setCellValue("Revenue", "B1", "Revenue");
    engine.setCellValue("Revenue", "A2", "West");
    engine.setCellFormula("Revenue", "B2", "SUM(B3:B5)");
    engine.setFreezePane("Revenue", 1, 0);
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Revenue",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 8,
            },
          },
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "summarizeCurrentSheet",
        },
      });

      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "summarizeCurrentSheet",
          title: "Summarize Current Sheet",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "inspect-current-sheet",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            title: "Current Sheet Summary",
            text: expect.stringContaining("Sheet: Revenue"),
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Started workflow: Summarize Current Sheet",
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Summarize Current Sheet",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Revenue",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("runs durable dependency trace workflows from the current selection context", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setCellFormula("Sheet1", "C1", "B1+1");
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B1",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
            },
          },
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "traceSelectionDependencies",
        },
      });

      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "traceSelectionDependencies",
          title: "Trace Selection Dependencies",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "trace-links",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            title: "Dependency Trace",
            text: expect.stringContaining("Root: Sheet1!B1"),
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Trace Selection Dependencies",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Sheet1",
                startAddress: "B1",
                endAddress: "B1",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("runs durable current-cell explanation workflows from the active selection", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setCellFormula("Sheet1", "C1", "B1+1");
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B1",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
            },
          },
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "explainSelectionCell",
        },
      });

      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "explainSelectionCell",
          title: "Explain Current Cell",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "explain-cell",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            title: "Current Cell",
            text: expect.stringContaining("Cell: Sheet1!B1"),
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Explain Current Cell",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Sheet1",
                startAddress: "B1",
                endAddress: "B1",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("runs durable workbook search workflows with query input", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Revenue");
    engine.setCellValue("Revenue", "A1", "Region");
    engine.setCellValue("Revenue", "B1", "Revenue");
    engine.setCellFormula("Revenue", "B2", "SUM(B3:B5)");
    engine.setCellValue("Revenue", "A2", "West");
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "searchWorkbookQuery",
          query: "revenue",
          limit: 5,
        },
      });

      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "searchWorkbookQuery",
          title: "Search Workbook",
          status: "completed",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "search-workbook",
              status: "completed",
            }),
          ]),
          artifact: expect.objectContaining({
            title: "Workbook Search",
            text: expect.stringContaining("Query: revenue"),
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Search Workbook",
            citations: expect.arrayContaining([
              expect.objectContaining({
                kind: "range",
                sheetName: "Revenue",
              }),
            ]),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("stages create-sheet preview bundles from durable workflows", async () => {
    const fakeCodex = new FakeCodexTransport();
    const getWorkbookHeadRevision = vi.fn(async () => 7);
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        getWorkbookHeadRevision,
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "createSheet",
          name: "Forecast",
        },
      });

      expect(getWorkbookHeadRevision).toHaveBeenCalledWith("doc-1");
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "createSheet",
          status: "completed",
          artifact: expect.objectContaining({
            title: "Create Sheet Preview",
          }),
        }),
      );
      expect(snapshot.pendingBundle).toEqual(
        expect.objectContaining({
          baseRevision: 7,
          turnId: expect.stringContaining("workflow:"),
          commands: [expect.objectContaining({ kind: "createSheet", name: "Forecast" })],
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: expect.stringContaining("Staged preview bundle"),
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Create Sheet",
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("stages rename-sheet preview bundles from durable workflows", async () => {
    const fakeCodex = new FakeCodexTransport();
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Revenue",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      const snapshot = await service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "renameCurrentSheet",
          name: "Forecast",
        },
      });

      expect(snapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "renameCurrentSheet",
          status: "completed",
          artifact: expect.objectContaining({
            title: "Rename Sheet Preview",
          }),
        }),
      );
      expect(snapshot.pendingBundle).toEqual(
        expect.objectContaining({
          commands: [
            expect.objectContaining({
              kind: "renameSheet",
              currentName: "Revenue",
              nextName: "Forecast",
            }),
          ],
        }),
      );
    } finally {
      await service.close();
    }
  });

  it("allows Codex dynamic tools to start durable workflows inside the active thread", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null };
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options;
          return fakeCodex;
        },
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-start-workflow",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "summarizeWorkbook",
        },
      });

      const snapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(result?.success).toBe(true);
      expect(snapshot.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            workflowTemplate: "summarizeWorkbook",
            status: "completed",
            steps: expect.arrayContaining([
              expect.objectContaining({
                stepId: "inspect-workbook",
                status: "completed",
              }),
            ]),
            artifact: expect.objectContaining({
              title: "Workbook Summary",
            }),
          }),
        ]),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Started workflow: Summarize Workbook",
          }),
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Summarize Workbook",
          }),
        ]),
      );
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
    } finally {
      await service.close();
    }
  });

  it("uses the request turn actor and context for shared-thread workflow starts", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null };
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.createSheet("Sheet2");
    engine.setCellValue("Sheet2", "C7", 99);
    let durableThreadState: WorkbookAgentThreadStateRecord | null = {
      documentId: "doc-1",
      threadId: "thr-shared",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "A1",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      entries: [],
      pendingBundle: null,
      updatedAtUnixMs: 100,
    };
    const upsertWorkbookWorkflowRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        async loadWorkbookAgentThreadState() {
          return durableThreadState ? structuredClone(durableThreadState) : null;
        },
        async saveWorkbookAgentThreadState(record: WorkbookAgentThreadStateRecord) {
          durableThreadState = structuredClone(record);
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options;
          return fakeCodex;
        },
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared",
          threadId: "thr-shared",
        },
      });

      const caseySnapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-casey",
          threadId: "thr-shared",
        },
      });

      await service.startTurn({
        documentId: "doc-1",
        sessionId: caseySnapshot.sessionId,
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          prompt: "Summarize my current sheet",
          context: {
            selection: {
              sheetName: "Sheet2",
              address: "C7",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      await service.updateContext({
        documentId: "doc-1",
        sessionId: caseySnapshot.sessionId,
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "A1",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      const result = await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-shared",
        turnId: "turn-1",
        callId: "call-start-workflow",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "summarizeCurrentSheet",
        },
      });

      const snapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: caseySnapshot.sessionId,
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
      });

      expect(result?.success).toBe(true);
      expect(snapshot.workflowRuns).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            workflowTemplate: "summarizeCurrentSheet",
            startedByUserId: "casey@example.com",
            summary: "Summarized Sheet2 with 1 populated cell and 0 tables.",
            artifact: expect.objectContaining({
              text: expect.stringContaining("Sheet: Sheet2"),
            }),
          }),
        ]),
      );
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      const startedByUserIds = upsertWorkbookWorkflowRun.mock.calls.map(
        (call) =>
          (
            call.at(1) as
              | {
                  startedByUserId?: string;
                }
              | undefined
          )?.startedByUserId ?? null,
      );
      expect(startedByUserIds).toEqual(["casey@example.com", "casey@example.com"]);
    } finally {
      await service.close();
    }
  });

  it("rejects starting a second workflow while one is still running", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);

    let releaseInspection!: () => void;
    const inspectBarrier = new Promise<void>((resolve) => {
      releaseInspection = () => {
        resolve();
      };
    });
    let resolveRunningPersisted!: () => void;
    const runningPersisted = new Promise<void>((resolve) => {
      resolveRunningPersisted = () => {
        resolve();
      };
    });
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          await inspectBarrier;
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        async upsertWorkbookWorkflowRun(_documentId, run) {
          if (run.status === "running") {
            resolveRunningPersisted();
          }
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const firstWorkflow = service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "summarizeWorkbook",
        },
      });

      await runningPersisted;

      await expect(
        service.startWorkflow({
          documentId: "doc-1",
          sessionId: "agent-session-1",
          session: {
            userID: "alex@example.com",
            roles: ["editor"],
          },
          body: {
            workflowTemplate: "describeRecentChanges",
          },
        }),
      ).rejects.toMatchObject({
        code: "WORKBOOK_AGENT_WORKFLOW_ALREADY_RUNNING",
        statusCode: 409,
      });

      releaseInspection();
      await firstWorkflow;
    } finally {
      releaseInspection();
      await service.close();
    }
  });

  it("rejects mutating workflows when a preview bundle is already staged", async () => {
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-existing",
            actorUserId: "alex@example.com",
            scope: "private",
            context: null,
            entries: [],
            pendingBundle: {
              id: "bundle-existing",
              documentId: "doc-1",
              threadId: "thr-existing",
              turnId: "turn-1",
              goalText: "Normalize the imported range",
              summary: "Normalize Sheet1!A1:A20",
              scope: "sheet",
              riskClass: "medium",
              approvalMode: "preview",
              baseRevision: 4,
              createdAtUnixMs: 100,
              context: null,
              commands: [
                {
                  kind: "formatRange",
                  range: {
                    sheetName: "Sheet1",
                    startAddress: "A1",
                    endAddress: "A20",
                  },
                  patch: {
                    font: {
                      bold: true,
                    },
                  },
                },
              ],
              affectedRanges: [
                {
                  sheetName: "Sheet1",
                  startAddress: "A1",
                  endAddress: "A20",
                  role: "target",
                },
              ],
              estimatedAffectedCells: 20,
            },
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          threadId: "thr-existing",
        },
      });

      await expect(
        service.startWorkflow({
          documentId: "doc-1",
          sessionId: snapshot.sessionId,
          session: {
            userID: "alex@example.com",
            roles: ["editor"],
          },
          body: {
            workflowTemplate: "createSheet",
            name: "Summary",
          },
        }),
      ).rejects.toMatchObject({
        code: "WORKBOOK_AGENT_PENDING_BUNDLE_EXISTS",
        statusCode: 409,
      });
    } finally {
      await service.close();
    }
  });

  it("cancels a running durable workflow without letting late completion overwrite it", async () => {
    const fakeCodex = new FakeCodexTransport();
    const engine = new SpreadsheetEngine({
      workbookName: "doc-1",
      replicaId: "server:test",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);

    let releaseInspection!: () => void;
    const inspectBarrier = new Promise<void>((resolve) => {
      releaseInspection = () => {
        resolve();
      };
    });
    let resolveRunningPersisted!: () => void;
    const runningPersisted = new Promise<void>((resolve) => {
      resolveRunningPersisted = () => {
        resolve();
      };
    });
    const upsertWorkbookWorkflowRun = vi.fn(async (_documentId: string, run) => {
      if (run.status === "running") {
        resolveRunningPersisted();
      }
    });

    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async inspectWorkbook<T>(
          _documentId: string,
          task: (runtime: WorkbookRuntime) => T | Promise<T>,
        ) {
          await inspectBarrier;
          const runtime: WorkbookRuntime = {
            documentId: "doc-1",
            engine,
            projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
              revision: 1,
              calculatedRevision: 1,
              ownerUserId: "alex@example.com",
              updatedBy: "alex@example.com",
              updatedAt: "2026-04-10T00:00:00.000Z",
            }),
            headRevision: 1,
            calculatedRevision: 1,
            ownerUserId: "alex@example.com",
          };
          return await task(runtime);
        },
        upsertWorkbookWorkflowRun,
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      const workflowPromise = service.startWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          workflowTemplate: "summarizeWorkbook",
        },
      });

      await runningPersisted;
      const runningSnapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });
      const runningRunId = runningSnapshot.workflowRuns[0]?.runId;
      if (!runningRunId) {
        throw new Error("Expected running workflow run id");
      }

      const cancelledSnapshot = await service.cancelWorkflow({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        runId: runningRunId,
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(cancelledSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "summarizeWorkbook",
          status: "cancelled",
          summary: "Cancelled workflow: Summarize Workbook",
          errorMessage: "Cancelled by alex@example.com.",
          steps: expect.arrayContaining([
            expect.objectContaining({
              stepId: "inspect-workbook",
              status: "cancelled",
            }),
          ]),
        }),
      );

      releaseInspection();
      const finalSnapshot = await workflowPromise;

      expect(finalSnapshot.workflowRuns[0]).toEqual(
        expect.objectContaining({
          workflowTemplate: "summarizeWorkbook",
          status: "cancelled",
          summary: "Cancelled workflow: Summarize Workbook",
          artifact: null,
        }),
      );
      expect(finalSnapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Started workflow: Summarize Workbook",
          }),
          expect.objectContaining({
            kind: "system",
            text: "Cancelled workflow: Summarize Workbook",
          }),
        ]),
      );
      expect(finalSnapshot.entries).not.toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Completed workflow: Summarize Workbook",
          }),
        ]),
      );
      expect(upsertWorkbookWorkflowRun).toHaveBeenCalledTimes(2);
      expect(upsertWorkbookWorkflowRun.mock.calls.map(([, run]) => run.status)).toEqual([
        "running",
        "cancelled",
      ]);
    } finally {
      releaseInspection();
      await service.close();
    }
  });

  it("uses nested app-server error messages instead of the generic fallback", async () => {
    const fakeCodex = new FakeCodexTransport();
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      fakeCodex.emit({
        method: "error",
        params: {
          error: {
            code: -32602,
            message: "thread/start.dynamicTools requires experimentalApi capability",
          },
        },
      });

      const snapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(snapshot.status).toBe("failed");
      expect(snapshot.lastError).toBe(
        "thread/start.dynamicTools requires experimentalApi capability",
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "thread/start.dynamicTools requires experimentalApi capability",
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("falls back to a stable runtime message when the app-server emits an empty error", async () => {
    const fakeCodex = new FakeCodexTransport();
    const service = createWorkbookAgentService(createZeroSyncStub(), {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
        },
      });

      fakeCodex.emit({
        method: "error",
        params: {},
      });

      const snapshot = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });

      expect(snapshot.status).toBe("failed");
      expect(snapshot.lastError).toBe("Workbook assistant runtime failed. Retry in a moment.");
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Workbook assistant runtime failed. Retry in a moment.",
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("persists the authoritative preview returned by apply and not the caller payload", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null };
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 7,
      preview: createPreviewSummary({
        cellDiffs: [
          {
            sheetName: "Sheet1",
            address: "B2",
            beforeInput: null,
            beforeFormula: null,
            afterInput: 42,
            afterFormula: null,
            changeKinds: ["input"],
          },
        ],
        effectSummary: {
          displayedCellDiffCount: 1,
          truncatedCellDiffs: false,
          inputChangeCount: 1,
          formulaChangeCount: 0,
          styleChangeCount: 0,
          numberFormatChangeCount: 0,
          structuralChangeCount: 0,
        },
      }),
    }));
    const appendWorkbookAgentRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options;
          return fakeCodex;
        },
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-1",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      });

      const pending = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      }).pendingBundle;

      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error("Expected a staged pending bundle");
      }

      const applied = await service.applyPendingBundle({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        bundleId: pending.id,
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        appliedBy: "user",
        preview: createPreviewSummary(),
      });

      const record = applied.executionRecords[0];
      if (!isWorkbookAgentExecutionRecord(record)) {
        throw new Error("Expected an execution record after apply");
      }
      expect(applyAgentCommandBundle).toHaveBeenCalled();
      expect(record.preview).toEqual(
        expect.objectContaining({
          cellDiffs: [
            expect.objectContaining({
              sheetName: "Sheet1",
              address: "B2",
            }),
          ],
        }),
      );
      expect(appendWorkbookAgentRun).toHaveBeenCalledWith(
        expect.objectContaining({
          preview: expect.objectContaining({
            effectSummary: expect.objectContaining({
              displayedCellDiffCount: 1,
              inputChangeCount: 1,
            }),
          }),
        }),
      );
      expect(applied.entries).toContainEqual(
        expect.objectContaining({
          kind: "system",
          text: "Applied preview bundle at revision r7: Write cells in Sheet1!B2",
          citations: [
            expect.objectContaining({
              kind: "range",
              sheetName: "Sheet1",
              startAddress: "B2",
              endAddress: "B2",
            }),
            expect.objectContaining({
              kind: "revision",
              revision: 7,
            }),
          ],
        }),
      );
    } finally {
      await service.close();
    }
  });

  it("keeps the pending bundle when authoritative apply rejects a stale preview", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null };
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle: vi.fn(async () => {
          throw createWorkbookAgentServiceError({
            code: "WORKBOOK_AGENT_PREVIEW_STALE",
            message:
              "Workbook changed after preview. Replay the plan to stage a fresh preview bundle.",
            statusCode: 409,
            retryable: true,
          });
        }),
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options;
          return fakeCodex;
        },
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-1",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      });

      const pending = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      }).pendingBundle;
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error("Expected a staged pending bundle");
      }

      await expect(
        service.applyPendingBundle({
          documentId: "doc-1",
          sessionId: "agent-session-1",
          bundleId: pending.id,
          session: {
            userID: "alex@example.com",
            roles: ["editor"],
          },
          appliedBy: "user",
          preview: createPreviewSummary(),
        }),
      ).rejects.toThrow("Replay the plan to stage a fresh preview bundle.");

      const afterFailure = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      });
      expect(isWorkbookAgentCommandBundle(afterFailure.pendingBundle)).toBe(true);
      if (!isWorkbookAgentCommandBundle(afterFailure.pendingBundle)) {
        throw new Error("Expected the pending bundle to remain staged");
      }
      expect(afterFailure.pendingBundle.id).toBe(pending.id);
      expect(afterFailure.executionRecords).toEqual([]);
    } finally {
      await service.close();
    }
  });

  it("applies a selected command subset and re-stages the remaining plan", async () => {
    const fakeCodex = new FakeCodexTransport();
    const capturedOptions: { current: CodexAppServerClientOptions | null } = { current: null };
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 7,
      preview: createPreviewSummary({
        cellDiffs: [
          {
            sheetName: "Sheet1",
            address: "C3",
            beforeInput: null,
            beforeFormula: null,
            afterInput: 2,
            afterFormula: null,
            changeKinds: ["input"],
          },
        ],
        effectSummary: {
          displayedCellDiffCount: 1,
          truncatedCellDiffs: false,
          inputChangeCount: 1,
          formulaChangeCount: 0,
          styleChangeCount: 0,
          numberFormatChangeCount: 0,
          structuralChangeCount: 0,
        },
      }),
    }));
    const appendWorkbookAgentRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
      }),
      {
        codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
          capturedOptions.current = options;
          return fakeCodex;
        },
      },
    );

    try {
      await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-1",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[1]],
        },
      });
      await capturedOptions.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-2",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "C3",
          values: [[2]],
        },
      });

      const pending = service.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      }).pendingBundle;
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error("Expected a staged pending bundle");
      }

      const applied = await service.applyPendingBundle({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        bundleId: pending.id,
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        appliedBy: "user",
        commandIndexes: [1],
        preview: createPreviewSummary({
          cellDiffs: [
            {
              sheetName: "Sheet1",
              address: "C3",
              beforeInput: null,
              beforeFormula: null,
              afterInput: 2,
              afterFormula: null,
              changeKinds: ["input"],
            },
          ],
          effectSummary: {
            displayedCellDiffCount: 1,
            truncatedCellDiffs: false,
            inputChangeCount: 1,
            formulaChangeCount: 0,
            styleChangeCount: 0,
            numberFormatChangeCount: 0,
            structuralChangeCount: 0,
          },
        }),
      });

      expect(applyAgentCommandBundle).toHaveBeenCalledWith(
        "doc-1",
        expect.objectContaining({
          id: pending.id,
          summary: "Write cells in Sheet1!C3",
          commands: [
            {
              kind: "writeRange",
              sheetName: "Sheet1",
              startAddress: "C3",
              values: [[2]],
            },
          ],
        }),
        expect.objectContaining({
          cellDiffs: [
            expect.objectContaining({
              address: "C3",
            }),
          ],
        }),
        expect.objectContaining({
          userID: "alex@example.com",
        }),
      );
      expect(applied.executionRecords[0]).toEqual(
        expect.objectContaining({
          bundleId: pending.id,
          acceptedScope: "partial",
          summary: "Write cells in Sheet1!C3",
          commands: [
            {
              kind: "writeRange",
              sheetName: "Sheet1",
              startAddress: "C3",
              values: [[2]],
            },
          ],
        }),
      );
      expect(applied.pendingBundle).toEqual(
        expect.objectContaining({
          baseRevision: 7,
          summary: "Write cells in Sheet1!B2",
          commands: [
            {
              kind: "writeRange",
              sheetName: "Sheet1",
              startAddress: "B2",
              values: [[1]],
            },
          ],
        }),
      );
      if (!isWorkbookAgentCommandBundle(applied.pendingBundle)) {
        throw new Error("Expected the remaining staged bundle to stay pending");
      }
      expect(applied.pendingBundle.id).not.toBe(pending.id);
      expect(appendWorkbookAgentRun).toHaveBeenCalledWith(
        expect.objectContaining({
          acceptedScope: "partial",
          summary: "Write cells in Sheet1!C3",
        }),
      );
    } finally {
      await service.close();
    }
  });

  it("recovers durable pending bundle state after the service restarts", async () => {
    let durableThreadState: WorkbookAgentThreadStateRecord | null = null;
    const fakeCodexA = new FakeCodexTransport();
    const capturedA: { current: CodexAppServerClientOptions | null } = { current: null };
    const zeroSync = createZeroSyncStub({
      async loadWorkbookAgentThreadState() {
        return durableThreadState ? structuredClone(durableThreadState) : null;
      },
      async saveWorkbookAgentThreadState(record) {
        durableThreadState = structuredClone(record);
      },
    });
    const serviceA = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (options: CodexAppServerClientOptions): CodexAppServerTransport => {
        capturedA.current = options;
        return fakeCodexA;
      },
    });

    try {
      await serviceA.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "A1",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      await capturedA.current?.handleDynamicToolCall({
        threadId: "thr-test",
        turnId: "turn-1",
        callId: "call-1",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      });

      const pending = serviceA.getSnapshot({
        documentId: "doc-1",
        sessionId: "agent-session-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
      }).pendingBundle;
      if (!isWorkbookAgentCommandBundle(pending)) {
        throw new Error("Expected a staged pending bundle before restart");
      }
      expect(durableThreadState).toEqual(
        expect.objectContaining({
          pendingBundle: expect.objectContaining({
            id: pending.id,
            summary: pending.summary,
          }),
        }),
      );
    } finally {
      await serviceA.close();
    }

    const fakeCodexB = new FakeCodexTransport();
    const serviceB = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodexB,
    });

    try {
      const resumed = await serviceB.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-2",
          threadId: "thr-test",
        },
      });

      expect(resumed.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            sheetName: "Sheet1",
            address: "A1",
          }),
        }),
      );
      expect(resumed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: expect.stringContaining("Write cells in Sheet1!B2"),
          }),
        ]),
      );
      expect(resumed.pendingBundle).toEqual(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-test",
          summary: "Write cells in Sheet1!B2",
        }),
      );
    } finally {
      await serviceB.close();
    }
  });

  it("falls back to durable thread state when live thread resume is unavailable", async () => {
    const fakeCodex = new FakeCodexTransport();
    fakeCodex.resumeError = new Error("codex resume unavailable");
    const durableThreadState: WorkbookAgentThreadStateRecord = {
      documentId: "doc-1",
      threadId: "thr-durable-only",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: {
        selection: {
          sheetName: "Sheet2",
          address: "C7",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      entries: [
        {
          id: "system-1",
          kind: "system",
          turnId: null,
          text: "Recovered durable shared thread history.",
          phase: null,
          toolName: null,
          toolStatus: null,
          argumentsText: null,
          outputText: null,
          success: null,
          citations: [],
        },
      ],
      pendingBundle: null,
      updatedAtUnixMs: 100,
    };
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return structuredClone(durableThreadState);
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          fakeCodex,
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-durable-only",
          threadId: "thr-durable-only",
        },
      });

      expect(fakeCodex.lastThreadResumeInput).toEqual(
        expect.objectContaining({
          threadId: "thr-durable-only",
        }),
      );
      expect(snapshot.threadId).toBe("thr-durable-only");
      expect(snapshot.scope).toBe("shared");
      expect(snapshot.status).toBe("failed");
      expect(snapshot.lastError).toContain("codex resume unavailable");
      expect(snapshot.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            sheetName: "Sheet2",
            address: "C7",
          }),
        }),
      );
      expect(snapshot.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: "Recovered durable shared thread history.",
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("allows collaborators to reuse a shared thread session while persisting the canonical owner row", async () => {
    const fakeCodex = new FakeCodexTransport();
    let durableThreadState: WorkbookAgentThreadStateRecord | null = {
      documentId: "doc-1",
      threadId: "thr-shared",
      actorUserId: "alex@example.com",
      scope: "shared",
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "A1",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      entries: [],
      pendingBundle: null,
      updatedAtUnixMs: 100,
    };
    const saveWorkbookAgentThreadState = vi.fn(async (record: WorkbookAgentThreadStateRecord) => {
      durableThreadState = structuredClone(record);
    });
    const zeroSync = createZeroSyncStub({
      async loadWorkbookAgentThreadState() {
        return durableThreadState ? structuredClone(durableThreadState) : null;
      },
      saveWorkbookAgentThreadState,
    });
    const service = createWorkbookAgentService(zeroSync, {
      codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
        fakeCodex,
    });

    try {
      const alexSnapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared",
          threadId: "thr-shared",
        },
      });

      const caseySnapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-casey",
          threadId: "thr-shared",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 20,
              colStart: 0,
              colEnd: 10,
            },
          },
        },
      });

      expect(alexSnapshot.scope).toBe("shared");
      expect(caseySnapshot.sessionId).toBe("agent-session-shared");
      expect(caseySnapshot.scope).toBe("shared");
      expect(caseySnapshot.context).toEqual(
        expect.objectContaining({
          selection: expect.objectContaining({
            address: "B2",
          }),
        }),
      );

      await service.startTurn({
        documentId: "doc-1",
        sessionId: caseySnapshot.sessionId,
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          prompt: "Review this shared thread",
        },
      });

      expect(saveWorkbookAgentThreadState).toHaveBeenLastCalledWith(
        expect.objectContaining({
          actorUserId: "alex@example.com",
          scope: "shared",
        }),
      );
    } finally {
      await service.close();
    }
  });

  it("loads shared execution history when a collaborator resumes a shared thread", async () => {
    const executionRecord = {
      id: "run-shared-1",
      bundleId: "bundle-shared-1",
      documentId: "doc-1",
      threadId: "thr-shared",
      turnId: "turn-1",
      actorUserId: "alex@example.com",
      goalText: "Normalize imported rows",
      planText: "Apply the shared cleanup plan",
      summary: "Write cells in Sheet1!B2",
      scope: "sheet" as const,
      riskClass: "medium" as const,
      approvalMode: "preview" as const,
      acceptedScope: "full" as const,
      appliedBy: "user" as const,
      baseRevision: 3,
      appliedRevision: 4,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 200,
      context: null,
      commands: [
        {
          kind: "writeRange" as const,
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[42]],
        },
      ],
      preview: null,
    };
    const listWorkbookAgentThreadRuns = vi.fn(async () => [executionRecord]);
    const listWorkbookAgentRuns = vi.fn(async () => []);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        listWorkbookAgentRuns,
        listWorkbookAgentThreadRuns,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-shared",
            actorUserId: "alex@example.com",
            scope: "shared",
            context: null,
            entries: [],
            pendingBundle: null,
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared",
          threadId: "thr-shared",
        },
      });

      expect(snapshot.scope).toBe("shared");
      expect(snapshot.executionRecords).toEqual([executionRecord]);
      expect(listWorkbookAgentThreadRuns).toHaveBeenCalledWith(
        "doc-1",
        "casey@example.com",
        "thr-shared",
      );
      expect(listWorkbookAgentRuns).not.toHaveBeenCalled();
    } finally {
      await service.close();
    }
  });

  it("requires the shared thread owner to apply medium/high-risk bundles", async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 5,
      preview: createPreviewSummary(),
    }));
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-shared",
            actorUserId: "alex@example.com",
            scope: "shared",
            context: null,
            entries: [],
            pendingBundle: {
              id: "bundle-shared-1",
              documentId: "doc-1",
              threadId: "thr-shared",
              turnId: "turn-1",
              goalText: "Build a workbook-wide summary",
              summary: "Create summary sheet and rewrite rollups",
              scope: "workbook",
              riskClass: "high",
              approvalMode: "explicit",
              baseRevision: 4,
              createdAtUnixMs: 100,
              context: null,
              commands: [
                {
                  kind: "createSheet",
                  name: "Summary",
                },
              ],
              affectedRanges: [],
              estimatedAffectedCells: 0,
            },
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared",
          threadId: "thr-shared",
        },
      });

      await expect(
        service.applyPendingBundle({
          documentId: "doc-1",
          sessionId: snapshot.sessionId,
          bundleId: "bundle-shared-1",
          session: {
            userID: "casey@example.com",
            roles: ["editor"],
          },
          appliedBy: "user",
          preview: createPreviewSummary(),
        }),
      ).rejects.toThrow(
        "Shared medium/high-risk workbook bundles must be applied by the thread owner.",
      );
      expect(applyAgentCommandBundle).not.toHaveBeenCalled();
    } finally {
      await service.close();
    }
  });

  it("still allows collaborators to apply low-risk shared bundles manually", async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 5,
      preview: createPreviewSummary(),
    }));
    const appendWorkbookAgentRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-shared",
            actorUserId: "alex@example.com",
            scope: "shared",
            context: null,
            entries: [],
            pendingBundle: {
              id: "bundle-shared-low",
              documentId: "doc-1",
              threadId: "thr-shared",
              turnId: "turn-1",
              goalText: "Fix one visible cell",
              summary: "Write cells in Sheet1!B2",
              scope: "selection",
              riskClass: "low",
              approvalMode: "auto",
              baseRevision: 4,
              createdAtUnixMs: 100,
              context: null,
              commands: [
                {
                  kind: "writeRange",
                  sheetName: "Sheet1",
                  startAddress: "B2",
                  values: [[42]],
                },
              ],
              affectedRanges: [
                {
                  sheetName: "Sheet1",
                  startAddress: "B2",
                  endAddress: "B2",
                  role: "target",
                },
              ],
              estimatedAffectedCells: 1,
            },
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared-low",
          threadId: "thr-shared",
        },
      });

      const applied = await service.applyPendingBundle({
        documentId: "doc-1",
        sessionId: snapshot.sessionId,
        bundleId: "bundle-shared-low",
        session: {
          userID: "casey@example.com",
          roles: ["editor"],
        },
        appliedBy: "user",
        preview: createPreviewSummary(),
      });

      expect(applyAgentCommandBundle).toHaveBeenCalled();
      expect(applied.pendingBundle).toBeNull();
    } finally {
      await service.close();
    }
  });

  it("requires owner approval before applying shared medium/high-risk bundles", async () => {
    const applyAgentCommandBundle = vi.fn(async () => ({
      revision: 6,
      preview: createPreviewSummary(),
    }));
    const appendWorkbookAgentRun = vi.fn(async () => undefined);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        applyAgentCommandBundle,
        appendWorkbookAgentRun,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-shared",
            actorUserId: "alex@example.com",
            scope: "shared",
            context: null,
            entries: [],
            pendingBundle: {
              id: "bundle-shared-review",
              documentId: "doc-1",
              threadId: "thr-shared",
              turnId: "turn-1",
              goalText: "Normalize the workbook",
              summary: "Normalize shared workbook structure",
              scope: "workbook",
              riskClass: "high",
              approvalMode: "explicit",
              baseRevision: 4,
              createdAtUnixMs: 100,
              context: null,
              commands: [
                {
                  kind: "createSheet",
                  name: "Summary",
                },
              ],
              affectedRanges: [],
              estimatedAffectedCells: 0,
              sharedReview: {
                ownerUserId: "alex@example.com",
                status: "pending",
                decidedByUserId: null,
                decidedAtUnixMs: null,
                recommendations: [],
              },
            },
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared-owner",
          threadId: "thr-shared",
        },
      });

      await expect(
        service.applyPendingBundle({
          documentId: "doc-1",
          sessionId: snapshot.sessionId,
          bundleId: "bundle-shared-review",
          session: {
            userID: "alex@example.com",
            roles: ["editor"],
          },
          appliedBy: "user",
          preview: createPreviewSummary(),
        }),
      ).rejects.toThrow(
        "Shared medium/high-risk workbook bundles must be approved by the thread owner before apply.",
      );

      const reviewed = await service.reviewPendingBundle({
        documentId: "doc-1",
        sessionId: snapshot.sessionId,
        bundleId: "bundle-shared-review",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          decision: "approved",
        },
      });

      expect(reviewed.pendingBundle).toEqual(
        expect.objectContaining({
          sharedReview: expect.objectContaining({
            status: "approved",
            decidedByUserId: "alex@example.com",
            recommendations: [],
          }),
        }),
      );

      const applied = await service.applyPendingBundle({
        documentId: "doc-1",
        sessionId: snapshot.sessionId,
        bundleId: "bundle-shared-review",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        appliedBy: "user",
        preview: createPreviewSummary(),
      });

      expect(applyAgentCommandBundle).toHaveBeenCalled();
      expect(appendWorkbookAgentRun).toHaveBeenCalled();
      expect(applied.pendingBundle).toBeNull();
    } finally {
      await service.close();
    }
  });

  it("records collaborator recommendations before owner approval", async () => {
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-shared",
            actorUserId: "alex@example.com",
            scope: "shared",
            context: null,
            entries: [],
            pendingBundle: {
              id: "bundle-shared-review",
              documentId: "doc-1",
              threadId: "thr-shared",
              turnId: "turn-1",
              goalText: "Normalize the workbook",
              summary: "Normalize shared workbook structure",
              scope: "workbook",
              riskClass: "high",
              approvalMode: "explicit",
              baseRevision: 4,
              createdAtUnixMs: 100,
              context: null,
              commands: [
                {
                  kind: "createSheet",
                  name: "Summary",
                },
              ],
              affectedRanges: [],
              estimatedAffectedCells: 0,
              sharedReview: {
                ownerUserId: "alex@example.com",
                status: "pending",
                decidedByUserId: null,
                decidedAtUnixMs: null,
                recommendations: [],
              },
            },
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "pat@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-shared-collab",
          threadId: "thr-shared",
        },
      });

      const reviewed = await service.reviewPendingBundle({
        documentId: "doc-1",
        sessionId: snapshot.sessionId,
        bundleId: "bundle-shared-review",
        session: {
          userID: "pat@example.com",
          roles: ["editor"],
        },
        body: {
          decision: "approved",
        },
      });

      expect(reviewed.pendingBundle).toEqual(
        expect.objectContaining({
          sharedReview: expect.objectContaining({
            status: "pending",
            decidedByUserId: null,
            recommendations: [
              expect.objectContaining({
                userId: "pat@example.com",
                decision: "approved",
              }),
            ],
          }),
        }),
      );
      expect(reviewed.entries).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: "system",
            text: expect.stringContaining("pat@example.com recommended approval"),
          }),
        ]),
      );
    } finally {
      await service.close();
    }
  });

  it("loads only thread-scoped execution history for private threads", async () => {
    const threadExecutionRecord = {
      id: "run-private-1",
      bundleId: "bundle-private-1",
      documentId: "doc-1",
      threadId: "thr-private",
      turnId: "turn-1",
      actorUserId: "alex@example.com",
      goalText: "Fix the selected range",
      planText: "Repair formulas in the active thread",
      summary: "Write formulas in Sheet1!C2:C5",
      scope: "selection" as const,
      riskClass: "low" as const,
      approvalMode: "auto" as const,
      acceptedScope: "full" as const,
      appliedBy: "auto" as const,
      baseRevision: 7,
      appliedRevision: 8,
      createdAtUnixMs: 100,
      appliedAtUnixMs: 110,
      context: null,
      commands: [
        {
          kind: "writeRange" as const,
          sheetName: "Sheet1",
          startAddress: "C2",
          values: [[42]],
        },
      ],
      preview: null,
    };
    const foreignExecutionRecord = {
      ...threadExecutionRecord,
      id: "run-foreign-1",
      bundleId: "bundle-foreign-1",
      threadId: "thr-other",
      goalText: "Foreign thread run",
      summary: "Should never hydrate into thr-private",
    };
    const listWorkbookAgentThreadRuns = vi.fn(async () => [threadExecutionRecord]);
    const listWorkbookAgentRuns = vi.fn(async () => [foreignExecutionRecord]);
    const service = createWorkbookAgentService(
      createZeroSyncStub({
        listWorkbookAgentRuns,
        listWorkbookAgentThreadRuns,
        async loadWorkbookAgentThreadState() {
          return {
            documentId: "doc-1",
            threadId: "thr-private",
            actorUserId: "alex@example.com",
            scope: "private",
            context: null,
            entries: [],
            pendingBundle: null,
            updatedAtUnixMs: 100,
          };
        },
      }),
      {
        codexClientFactory: (_options: CodexAppServerClientOptions): CodexAppServerTransport =>
          new FakeCodexTransport(),
      },
    );

    try {
      const snapshot = await service.createSession({
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        body: {
          sessionId: "agent-session-private",
          threadId: "thr-private",
        },
      });

      expect(snapshot.scope).toBe("private");
      expect(snapshot.executionRecords).toEqual([threadExecutionRecord]);
      expect(listWorkbookAgentThreadRuns).toHaveBeenCalledWith(
        "doc-1",
        "alex@example.com",
        "thr-private",
      );
      expect(listWorkbookAgentRuns).not.toHaveBeenCalled();
    } finally {
      await service.close();
    }
  });
});
