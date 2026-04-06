import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { WorkbookAgentSessionSnapshot } from "@bilig/contracts";
import type { ZeroSyncService } from "../zero/service.js";
import { createWorkbookAgentServiceError } from "../workbook-agent-errors.js";
import { createSyncServer } from "./sync-server.js";

type TestServer = Awaited<ReturnType<typeof startHttpServer>>;

async function startHttpServer(
  handler: (request: IncomingMessage, response: ServerResponse) => void,
) {
  const server = createServer(handler);
  await new Promise<void>((resolve, reject) => {
    server.listen(0, "127.0.0.1", () => resolve());
    server.once("error", reject);
  });
  const address = server.address();
  if (!address || typeof address === "string") {
    throw new Error("Expected TCP test server address");
  }
  return {
    server,
    origin: `http://127.0.0.1:${String(address.port)}`,
    async close() {
      await new Promise<void>((resolve, reject) => {
        server.close((error) => {
          if (error) {
            reject(error);
            return;
          }
          resolve();
        });
      });
    },
  };
}

const upstreamServers: TestServer[] = [];

afterEach(async () => {
  delete process.env["BILIG_ZERO_PROXY_UPSTREAM"];
  await Promise.all(upstreamServers.splice(0).map((server) => server.close()));
});

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
    async applyServerMutator() {
      throw new Error("not used");
    },
    async applyAgentCommandBundle() {
      throw new Error("not used");
    },
    async listWorkbookAgentRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {
      throw new Error("not used");
    },
    async getWorkbookHeadRevision() {
      return 1;
    },
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
    async createWorkbookScenario() {
      throw new Error("not used");
    },
    async deleteWorkbookScenario() {
      throw new Error("not used");
    },
    ...overrides,
  };
}

function createAgentSessionSnapshot(
  overrides: Partial<WorkbookAgentSessionSnapshot> = {},
): WorkbookAgentSessionSnapshot {
  return {
    sessionId: "agent-session-1",
    documentId: "doc-1",
    threadId: "thr-1",
    status: "idle",
    activeTurnId: null,
    lastError: null,
    context: {
      selection: {
        sheetName: "Sheet1",
        address: "A1",
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    },
    entries: [],
    pendingBundle: null,
    executionRecords: [],
    ...overrides,
  };
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

describe("sync-server zero keepalive", () => {
  it("proxies a healthy keepalive response without using the generic zero proxy route", async () => {
    const upstream = await startHttpServer((request, response) => {
      expect(request.url).toBe("/keepalive");
      response.writeHead(200, { "content-type": "text/plain; charset=utf-8" });
      response.end("ok");
    });
    upstreamServers.push(upstream);
    process.env["BILIG_ZERO_PROXY_UPSTREAM"] = upstream.origin;

    const { app } = createSyncServer({ logger: false });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/zero/keepalive",
      });

      expect(response.statusCode).toBe(200);
      expect(response.headers["cache-control"]).toBe("no-store");
      expect(response.body).toBe("ok");
    } finally {
      await app.close();
    }
  });

  it("returns 503 when the upstream resets the keepalive connection", async () => {
    const upstream = await startHttpServer((request) => {
      expect(request.url).toBe("/keepalive");
      request.socket.destroy();
    });
    upstreamServers.push(upstream);
    process.env["BILIG_ZERO_PROXY_UPSTREAM"] = upstream.origin;

    const { app } = createSyncServer({ logger: false });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/zero/keepalive",
      });

      expect(response.statusCode).toBe(503);
      expect(response.json()).toEqual({
        error: "ZERO_CACHE_UNAVAILABLE",
        message: "Zero cache keepalive probe failed",
        retryable: true,
      });
    } finally {
      await app.close();
    }
  });
});

describe("sync-server authoritative events", () => {
  it("returns authoritative workbook events from the zero sync service", async () => {
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        async loadAuthoritativeEvents(documentId, afterRevision) {
          expect(documentId).toBe("doc-1");
          expect(afterRevision).toBe(4);
          return {
            afterRevision,
            headRevision: 6,
            calculatedRevision: 6,
            events: [
              {
                revision: 5,
                clientMutationId: "doc-1:pending:5",
                payload: {
                  kind: "setCellValue",
                  sheetName: "Sheet1",
                  address: "A1",
                  value: 42,
                },
              },
            ],
          };
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/events?afterRevision=4",
      });

      expect(response.statusCode).toBe(200);
      expect(response.headers["cache-control"]).toBe("no-store");
      expect(response.json()).toEqual({
        afterRevision: 4,
        headRevision: 6,
        calculatedRevision: 6,
        events: [
          {
            revision: 5,
            clientMutationId: "doc-1:pending:5",
            payload: {
              kind: "setCellValue",
              sheetName: "Sheet1",
              address: "A1",
              value: 42,
            },
          },
        ],
      });
    } finally {
      await app.close();
    }
  });

  it("rejects invalid afterRevision values", async () => {
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        async loadAuthoritativeEvents() {
          throw new Error("not used");
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/events?afterRevision=nope",
      });

      expect(response.statusCode).toBe(400);
      expect(response.json()).toEqual({
        error: "INVALID_AFTER_REVISION",
        message: "afterRevision must be a non-negative integer",
        retryable: false,
      });
    } finally {
      await app.close();
    }
  });
});

describe("sync-server workbook agent", () => {
  it("creates workbook agent sessions through the monolith route", async () => {
    const createSession = vi.fn(async () => createAgentSessionSnapshot());

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: {
        enabled: true,
        createSession,
        async updateContext() {
          throw new Error("not used");
        },
        async startTurn() {
          throw new Error("not used");
        },
        async interruptTurn() {
          throw new Error("not used");
        },
        async applyPendingBundle() {
          throw new Error("not used");
        },
        async dismissPendingBundle() {
          throw new Error("not used");
        },
        async replayExecutionRecord() {
          throw new Error("not used");
        },
        getSnapshot() {
          throw new Error("not used");
        },
        subscribe() {
          return () => {};
        },
        async close() {},
      },
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/agent/sessions",
        payload: {
          sessionId: "agent-session-1",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "A1",
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

      expect(response.statusCode).toBe(200);
      expect(response.headers["cache-control"]).toBe("no-store");
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: expect.objectContaining({
            sessionId: "agent-session-1",
          }),
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          sessionId: "agent-session-1",
          threadId: "thr-1",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("applies staged workbook bundles through the monolith route", async () => {
    const applyPendingBundle = vi.fn(async () =>
      createAgentSessionSnapshot({
        pendingBundle: null,
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: {
        enabled: true,
        async createSession() {
          throw new Error("not used");
        },
        async updateContext() {
          throw new Error("not used");
        },
        async startTurn() {
          throw new Error("not used");
        },
        async interruptTurn() {
          throw new Error("not used");
        },
        applyPendingBundle,
        async dismissPendingBundle() {
          throw new Error("not used");
        },
        async replayExecutionRecord() {
          throw new Error("not used");
        },
        getSnapshot() {
          throw new Error("not used");
        },
        subscribe() {
          return () => {};
        },
        async close() {},
      },
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/agent/sessions/agent-session-1/bundles/bundle-1/apply",
        payload: {
          commandIndexes: [1],
          preview: createPreviewSummary(),
        },
      });

      expect(response.statusCode).toBe(200);
      expect(applyPendingBundle).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          sessionId: "agent-session-1",
          bundleId: "bundle-1",
          appliedBy: "user",
          commandIndexes: [1],
          preview: createPreviewSummary(),
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("returns a structured conflict envelope when agent apply rejects a stale preview", async () => {
    const applyPendingBundle = vi.fn(async () => {
      throw createWorkbookAgentServiceError({
        code: "WORKBOOK_AGENT_PREVIEW_STALE",
        message: "Workbook changed after preview. Replay the plan to stage a fresh preview bundle.",
        statusCode: 409,
        retryable: true,
      });
    });

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: {
        enabled: true,
        async createSession() {
          throw new Error("not used");
        },
        async updateContext() {
          throw new Error("not used");
        },
        async startTurn() {
          throw new Error("not used");
        },
        async interruptTurn() {
          throw new Error("not used");
        },
        applyPendingBundle,
        async dismissPendingBundle() {
          throw new Error("not used");
        },
        async replayExecutionRecord() {
          throw new Error("not used");
        },
        getSnapshot() {
          throw new Error("not used");
        },
        subscribe() {
          return () => {};
        },
        async close() {},
      },
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/agent/sessions/agent-session-1/bundles/bundle-1/apply",
        payload: {
          preview: createPreviewSummary(),
        },
      });

      expect(response.statusCode).toBe(409);
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: "WORKBOOK_AGENT_PREVIEW_STALE",
          message:
            "Workbook changed after preview. Replay the plan to stage a fresh preview bundle.",
          retryable: true,
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("dismisses staged workbook bundles through the monolith route", async () => {
    const dismissPendingBundle = vi.fn(async () => createAgentSessionSnapshot());

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: {
        enabled: true,
        async createSession() {
          throw new Error("not used");
        },
        async updateContext() {
          throw new Error("not used");
        },
        async startTurn() {
          throw new Error("not used");
        },
        async interruptTurn() {
          throw new Error("not used");
        },
        async applyPendingBundle() {
          throw new Error("not used");
        },
        dismissPendingBundle,
        async replayExecutionRecord() {
          throw new Error("not used");
        },
        getSnapshot() {
          throw new Error("not used");
        },
        subscribe() {
          return () => {};
        },
        async close() {},
      },
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/agent/sessions/agent-session-1/bundles/bundle-1/dismiss",
      });

      expect(response.statusCode).toBe(200);
      expect(dismissPendingBundle).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          sessionId: "agent-session-1",
          bundleId: "bundle-1",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("replays prior execution records through the monolith route", async () => {
    const replayExecutionRecord = vi.fn(async () =>
      createAgentSessionSnapshot({
        pendingBundle: {
          id: "bundle-replay-1",
          documentId: "doc-1",
          threadId: "thr-1",
          turnId: "replay:run-1:10",
          goalText: "Reapply formatting",
          summary: "Format Sheet1!A1",
          scope: "selection",
          riskClass: "low",
          approvalMode: "auto",
          baseRevision: 4,
          createdAtUnixMs: 10,
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "A1",
            },
            viewport: {
              rowStart: 0,
              rowEnd: 10,
              colStart: 0,
              colEnd: 5,
            },
          },
          commands: [
            {
              kind: "formatRange",
              range: {
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "A1",
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
              endAddress: "A1",
              role: "target",
            },
          ],
          estimatedAffectedCells: 1,
        },
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: {
        enabled: true,
        async createSession() {
          throw new Error("not used");
        },
        async updateContext() {
          throw new Error("not used");
        },
        async startTurn() {
          throw new Error("not used");
        },
        async interruptTurn() {
          throw new Error("not used");
        },
        async applyPendingBundle() {
          throw new Error("not used");
        },
        async dismissPendingBundle() {
          throw new Error("not used");
        },
        replayExecutionRecord,
        getSnapshot() {
          throw new Error("not used");
        },
        subscribe() {
          return () => {};
        },
        async close() {},
      },
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/agent/sessions/agent-session-1/runs/run-1/replay",
      });

      expect(response.statusCode).toBe(200);
      expect(replayExecutionRecord).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          sessionId: "agent-session-1",
          recordId: "run-1",
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          pendingBundle: expect.objectContaining({
            id: "bundle-replay-1",
          }),
        }),
      );
    } finally {
      await app.close();
    }
  });
});

describe("sync-server workbook scenarios", () => {
  it("creates scenario branches through the monolith route", async () => {
    const createWorkbookScenario = vi.fn(async () => ({
      documentId: "scenario:new",
      workbookId: "doc-1",
      ownerUserId: "alex@example.com",
      name: "What-if plan",
      baseRevision: 12,
      sheetId: 3,
      sheetName: "Revenue",
      address: "D12",
      viewport: {
        rowStart: 4,
        rowEnd: 22,
        colStart: 2,
        colEnd: 10,
      },
      createdAt: 1_775_456_000_000,
      updatedAt: 1_775_456_000_000,
      browserUrl: "http://127.0.0.1:4321/?document=scenario%3Anew&sheet=Revenue&cell=D12",
    }));
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        createWorkbookScenario,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/scenarios",
        payload: {
          name: "What-if plan",
          sheetName: "Revenue",
          address: "D12",
          viewport: {
            rowStart: 4,
            rowEnd: 22,
            colStart: 2,
            colEnd: 10,
          },
        },
      });

      expect(response.statusCode).toBe(200);
      expect(createWorkbookScenario).toHaveBeenCalledWith(
        {
          workbookId: "doc-1",
          name: "What-if plan",
          sheetName: "Revenue",
          address: "D12",
          viewport: {
            rowStart: 4,
            rowEnd: 22,
            colStart: 2,
            colEnd: 10,
          },
        },
        expect.objectContaining({
          userID: expect.any(String),
        }),
      );
      expect(response.json()).toMatchObject({
        documentId: "scenario:new",
        workbookId: "doc-1",
        name: "What-if plan",
      });
    } finally {
      await app.close();
    }
  });

  it("deletes scenario branches through the monolith route", async () => {
    const deleteWorkbookScenario = vi.fn(async () => undefined);
    const { app } = createSyncServer({
      logger: false,
      zeroSyncService: createZeroSyncStub({
        deleteWorkbookScenario,
      }),
    });

    try {
      const response = await app.inject({
        method: "DELETE",
        url: "/v2/documents/doc-1/scenarios/scenario%3Adelete",
      });

      expect(response.statusCode).toBe(200);
      expect(deleteWorkbookScenario).toHaveBeenCalledWith(
        {
          workbookId: "doc-1",
          documentId: "scenario:delete",
        },
        expect.objectContaining({
          userID: expect.any(String),
        }),
      );
      expect(response.json()).toEqual({ ok: true });
    } finally {
      await app.close();
    }
  });
});
