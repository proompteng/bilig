import {
  isWorkbookAgentCommandBundle,
  isWorkbookAgentExecutionRecord,
  type CodexServerNotification,
  type CodexTurn,
} from "@bilig/agent-api";
import { describe, expect, it, vi } from "vitest";
import type { ZeroSyncService } from "../zero/service.js";
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
    async inspectWorkbook() {
      throw new Error("not used");
    },
    async applyServerMutator() {},
    async applyAgentCommandBundle() {
      return { revision: 2, preview: createPreviewSummary() };
    },
    async listWorkbookAgentRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {},
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

      expect(capturedOptions.current?.args).toEqual(["app-server"]);
      expect(fakeCodex.lastThreadStartInput?.dynamicTools.map((tool) => tool.name)).toEqual(
        expect.arrayContaining([
          "bilig.read_selection",
          "bilig.read_visible_range",
          "bilig.inspect_cell",
          "bilig.find_formula_issues",
          "bilig.search_workbook",
          "bilig.trace_dependencies",
          "bilig.read_range",
          "bilig.write_range",
        ]),
      );
      expect(fakeCodex.lastThreadStartInput?.baseInstructions).toContain("local workbook skills");
      expect(fakeCodex.lastThreadStartInput?.developerInstructions).toContain(
        "bilig.search_workbook",
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
      const unsubscribe = service.subscribe("agent-session-1", (event) => {
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
        tool: "bilig.write_range",
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
        tool: "bilig.write_range",
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
        tool: "bilig.write_range",
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
        tool: "bilig.write_range",
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
});
