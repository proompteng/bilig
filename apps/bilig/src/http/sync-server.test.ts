import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { WorkbookAgentThreadSnapshot } from "@bilig/contracts";
import { toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from "@bilig/agent-api";
import { Effect } from "effect";
import type { DocumentControlService } from "@bilig/runtime-kernel";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookAgentService } from "../codex-app/workbook-agent-service.js";
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
    async inspectWorkbook<T>(_documentId: string, _task: (runtime: never) => T | Promise<T>) {
      throw new Error("not used");
    },
    async applyServerMutator() {
      throw new Error("not used");
    },
    async applyAgentCommandBundle() {
      throw new Error("not used");
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
    async appendWorkbookAgentRun() {
      throw new Error("not used");
    },
    async listWorkbookAgentThreadSummaries() {
      return [];
    },
    async loadWorkbookAgentThreadState() {
      return null;
    },
    async saveWorkbookAgentThreadState() {
      throw new Error("not used");
    },
    async listWorkbookThreadWorkflowRuns() {
      return [];
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error("not used");
    },
    async getWorkbookHeadRevision() {
      return 1;
    },
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
    ...overrides,
  };
}

function createWorkbookAgentServiceStub(
  overrides: Partial<WorkbookAgentService> = {},
): WorkbookAgentService {
  return {
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
    async startWorkflow() {
      throw new Error("not used");
    },
    async cancelWorkflow() {
      throw new Error("not used");
    },
    async interruptTurn() {
      throw new Error("not used");
    },
    async applyPendingBundle() {
      throw new Error("not used");
    },
    async reviewPendingBundle() {
      throw new Error("not used");
    },
    async dismissPendingBundle() {
      throw new Error("not used");
    },
    async replayExecutionRecord() {
      throw new Error("not used");
    },
    async listThreads() {
      return [];
    },
    getObservabilitySnapshot() {
      return {
        enabled: true,
        generatedAtUnixMs: 1,
        featureFlags: {
          sharedThreadsEnabled: true,
          workflowRunnerEnabled: true,
          autoApplyLowRiskEnabled: true,
          formulaWorkflowFamilyEnabled: true,
          formattingWorkflowFamilyEnabled: true,
          importWorkflowFamilyEnabled: true,
          rollupWorkflowFamilyEnabled: true,
          structuralWorkflowFamilyEnabled: true,
          allowlistedUserCount: 0,
          allowlistedDocumentCount: 0,
        },
        sessions: {
          sessionCount: 0,
          subscriberThreadCount: 0,
          subscriberCount: 0,
          activeTurnCount: 0,
          runningWorkflowCount: 0,
          pendingBundleCount: 0,
          sharedPendingReviewCount: 0,
        },
        pool: {
          slotCount: 0,
          boundThreadCount: 0,
          activeTurnCount: 0,
          queuedTurnCount: 0,
          maxClients: 0,
          maxConcurrentTurnsPerClient: 0,
          maxQueuedTurnsPerClient: 0,
        },
        counters: {
          turnBackpressureCount: 0,
          workflowStartedCount: 0,
          workflowCompletedCount: 0,
          workflowFailedCount: 0,
          workflowCancelledCount: 0,
          sharedReviewApprovedCount: 0,
          sharedReviewRejectedCount: 0,
          sharedRecommendationApprovedCount: 0,
          sharedRecommendationRejectedCount: 0,
        },
      };
    },
    getSnapshot() {
      throw new Error("not used");
    },
    subscribe() {
      return () => {};
    },
    async close() {},
    ...overrides,
  };
}

function createDocumentServiceStub(
  overrides: Partial<DocumentControlService> = {},
): DocumentControlService {
  return {
    attachBrowser() {
      return Effect.sync(() => {
        throw new Error("not used");
      });
    },
    openBrowserSession() {
      return Effect.sync(() => {
        throw new Error("not used");
      });
    },
    handleSyncFrame() {
      return Effect.sync(() => {
        throw new Error("not used");
      });
    },
    handleAgentFrame() {
      return Effect.sync(() => {
        throw new Error("not used");
      });
    },
    getDocumentState() {
      return Effect.sync(() => {
        throw new Error("not used");
      });
    },
    getLatestSnapshot() {
      return Effect.succeed(null);
    },
    ...overrides,
  };
}

function createAgentSessionSnapshot(
  overrides: Partial<WorkbookAgentThreadSnapshot> = {},
): WorkbookAgentThreadSnapshot {
  return {
    documentId: "doc-1",
    threadId: "thr-1",
    executionPolicy: "autoApplyAll",
    scope: "private",
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
    reviewQueueItems: [],
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  };
}

function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? "ownerReview" : "manual",
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  });
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

describe("sync-server cross-origin isolation", () => {
  it("serves runtime responses with the headers required for SharedArrayBuffer-backed OPFS", async () => {
    const { app } = createSyncServer({ logger: false });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/runtime-config.json",
      });

      expect(response.statusCode).toBe(200);
      expect(response.headers["cross-origin-opener-policy"]).toBe("same-origin");
      expect(response.headers["cross-origin-embedder-policy"]).toBe("require-corp");
      expect(response.headers["origin-agent-cluster"]).toBe("?1");
    } finally {
      await app.close();
    }
  });
});

describe("sync-server snapshots", () => {
  it("returns 204 when no latest snapshot exists", async () => {
    const { app } = createSyncServer({
      logger: false,
      documentService: createDocumentServiceStub({
        getLatestSnapshot(documentId: string) {
          expect(documentId).toBe("doc-1");
          return Effect.succeed(null);
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/snapshot/latest",
      });

      expect(response.statusCode).toBe(204);
      expect(response.body).toBe("");
      expect(response.headers["content-type"]).toBeUndefined();
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
  it("lists durable workbook chat threads through the public route", async () => {
    const listThreads = vi.fn(async () => [
      {
        threadId: "thr-2",
        scope: "shared" as const,
        ownerUserId: "alex@example.com",
        updatedAtUnixMs: 200,
        entryCount: 3,
        reviewQueueItemCount: 0,
        latestEntryText: "Applied shared cleanup at revision r7",
      },
      {
        threadId: "thr-1",
        scope: "private" as const,
        ownerUserId: "alex@example.com",
        updatedAtUnixMs: 100,
        entryCount: 1,
        reviewQueueItemCount: 1,
        latestEntryText: "Preview bundle staged",
      },
    ]);

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        listThreads,
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/chat/threads",
      });

      expect(response.statusCode).toBe(200);
      expect(response.headers["cache-control"]).toBe("no-store");
      expect(listThreads).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
        }),
      );
      expect(response.json()).toEqual([
        {
          threadId: "thr-2",
          scope: "shared",
          ownerUserId: "alex@example.com",
          updatedAtUnixMs: 200,
          entryCount: 3,
          reviewQueueItemCount: 0,
          latestEntryText: "Applied shared cleanup at revision r7",
        },
        {
          threadId: "thr-1",
          scope: "private",
          ownerUserId: "alex@example.com",
          updatedAtUnixMs: 100,
          entryCount: 1,
          reviewQueueItemCount: 1,
          latestEntryText: "Preview bundle staged",
        },
      ]);
    } finally {
      await app.close();
    }
  });

  it("creates or resumes workbook chat threads through the public route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-shared",
        scope: "shared",
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads",
        payload: {
          threadId: "thr-shared",
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
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
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: expect.objectContaining({
            threadId: "thr-shared",
          }),
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: "thr-shared",
          scope: "shared",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("loads workbook chat thread snapshots through a thread-specific route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-shared",
        scope: "shared",
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/chat/threads/thr-shared",
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-shared",
          },
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: "thr-shared",
          scope: "shared",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("starts workbook chat turns through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const startTurn = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        status: "inProgress",
        activeTurnId: "turn-1",
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startTurn,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/turns",
        payload: {
          prompt: "Summarize this thread",
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
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: expect.objectContaining({
            threadId: "thr-2",
          }),
        }),
      );
      expect(startTurn).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          body: expect.objectContaining({
            prompt: "Summarize this thread",
          }),
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("starts workbook chat workflows through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const startWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        workflowRuns: [
          {
            runId: "wf-2",
            threadId: "thr-2",
            startedByUserId: "alex@example.com",
            workflowTemplate: "describeRecentChanges",
            title: "Describe Recent Changes",
            summary: "Summarized 3 recent workbook changes.",
            status: "completed" as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 3,
            completedAtUnixMs: 3,
            errorMessage: null,
            steps: [
              {
                stepId: "load-revisions",
                label: "Load durable revisions",
                status: "completed" as const,
                summary: "Loaded 3 durable workbook revisions.",
                updatedAtUnixMs: 2,
              },
              {
                stepId: "draft-change-report",
                label: "Draft change report",
                status: "completed" as const,
                summary: "Prepared the durable recent change report for the thread.",
                updatedAtUnixMs: 3,
              },
            ],
            artifact: {
              kind: "markdown" as const,
              title: "Recent Changes",
              text: "## Recent Changes",
            },
          },
        ],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startWorkflow,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/workflows",
        payload: {
          workflowTemplate: "describeRecentChanges",
        },
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(startWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          body: {
            workflowTemplate: "describeRecentChanges",
          },
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: "thr-2",
          workflowRuns: [
            expect.objectContaining({
              runId: "wf-2",
              workflowTemplate: "describeRecentChanges",
            }),
          ],
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("cancels workbook chat workflows through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const cancelWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        workflowRuns: [
          {
            runId: "wf-running-2",
            threadId: "thr-2",
            startedByUserId: "alex@example.com",
            workflowTemplate: "describeRecentChanges",
            title: "Describe Recent Changes",
            summary: "Cancelled workflow: Describe Recent Changes",
            status: "cancelled" as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 4,
            completedAtUnixMs: 4,
            errorMessage: "Cancelled by alex@example.com.",
            steps: [
              {
                stepId: "load-revisions",
                label: "Load durable revisions",
                status: "cancelled" as const,
                summary: "Workflow cancelled before this step completed.",
                updatedAtUnixMs: 4,
              },
            ],
            artifact: null,
          },
        ],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        cancelWorkflow,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/workflows/wf-running-2/cancel",
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(cancelWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          runId: "wf-running-2",
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          threadId: "thr-2",
          workflowRuns: [
            expect.objectContaining({
              runId: "wf-running-2",
              status: "cancelled",
            }),
          ],
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("passes query input through workbook search workflows", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-search",
      }),
    );
    const startWorkflow = vi.fn(async () =>
      createAgentSessionSnapshot({
        workflowRuns: [
          {
            runId: "wf-search-1",
            threadId: "thr-1",
            startedByUserId: "alex@example.com",
            workflowTemplate: "searchWorkbookQuery",
            title: "Search Workbook",
            summary: 'Found 1 workbook match for "revenue".',
            status: "completed" as const,
            createdAtUnixMs: 1,
            updatedAtUnixMs: 2,
            completedAtUnixMs: 2,
            errorMessage: null,
            steps: [
              {
                stepId: "search-workbook",
                label: "Search workbook",
                status: "completed" as const,
                summary:
                  'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 1 match.',
                updatedAtUnixMs: 1,
              },
              {
                stepId: "draft-search-report",
                label: "Draft search report",
                status: "completed" as const,
                summary: "Prepared the durable workbook search report for the thread.",
                updatedAtUnixMs: 2,
              },
            ],
            artifact: {
              kind: "markdown" as const,
              title: "Workbook Search",
              text: "## Workbook Search",
            },
          },
        ],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        startWorkflow,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-search/workflows",
        payload: {
          workflowTemplate: "searchWorkbookQuery",
          query: "revenue",
          limit: 5,
        },
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-search",
          },
        }),
      );
      expect(startWorkflow).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-search",
          body: {
            workflowTemplate: "searchWorkbookQuery",
            query: "revenue",
            limit: 5,
          },
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("reviews workbook agent bundles through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const reviewPendingBundle = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        reviewQueueItems: [
          createReviewQueueItem({
            id: "bundle-1",
            documentId: "doc-1",
            threadId: "thr-2",
            turnId: "turn-1",
            goalText: "Normalize shared workbook",
            summary: "Normalize shared workbook",
            scope: "workbook",
            riskClass: "high",
            approvalMode: "explicit",
            baseRevision: 4,
            createdAtUnixMs: 10,
            context: null,
            commands: [],
            affectedRanges: [],
            estimatedAffectedCells: 0,
            sharedReview: {
              ownerUserId: "alex@example.com",
              status: "approved",
              decidedByUserId: "alex@example.com",
              decidedAtUnixMs: 12,
              recommendations: [],
            },
          }),
        ],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        reviewPendingBundle,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/bundles/bundle-1/review",
        payload: {
          decision: "approved",
        },
      });

      expect(response.statusCode).toBe(200);
      expect(reviewPendingBundle).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          bundleId: "bundle-1",
          body: {
            decision: "approved",
          },
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [
            expect.objectContaining({
              reviewMode: "ownerReview",
              status: "approved",
              decidedByUserId: "alex@example.com",
            }),
          ],
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("updates workbook agent context through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const updateContext = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        updateContext,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/context",
        payload: {
          context: {
            selection: {
              sheetName: "Sheet1",
              address: "B2",
            },
            viewport: {
              rowStart: 1,
              rowEnd: 11,
              colStart: 1,
              colEnd: 6,
            },
          },
        },
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(updateContext).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          body: {
            context: {
              selection: {
                sheetName: "Sheet1",
                address: "B2",
              },
              viewport: {
                rowStart: 1,
                rowEnd: 11,
                colStart: 1,
                colEnd: 6,
              },
            },
          },
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("interrupts workbook agent turns through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const interruptTurn = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        status: "idle",
        activeTurnId: null,
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        interruptTurn,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/interrupt",
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(interruptTurn).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("applies staged workbook bundles through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const applyPendingBundle = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        reviewQueueItems: [],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyPendingBundle,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/bundles/bundle-1/apply",
        payload: {
          commandIndexes: [1],
          preview: createPreviewSummary(),
        },
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(applyPendingBundle).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
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
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-stale",
      }),
    );
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
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        applyPendingBundle,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-stale/bundles/bundle-1/apply",
        payload: {
          preview: createPreviewSummary(),
        },
      });

      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-stale",
          },
        }),
      );
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

  it("dismisses staged workbook bundles through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const dismissPendingBundle = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        dismissPendingBundle,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/bundles/bundle-1/dismiss",
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(dismissPendingBundle).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          bundleId: "bundle-1",
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("replays prior execution records through the public thread route", async () => {
    const createSession = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
      }),
    );
    const replayExecutionRecord = vi.fn(async () =>
      createAgentSessionSnapshot({
        threadId: "thr-2",
        reviewQueueItems: [
          createReviewQueueItem({
            id: "bundle-replay-1",
            documentId: "doc-1",
            threadId: "thr-2",
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
            sharedReview: null,
          }),
        ],
      }),
    );

    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        createSession,
        replayExecutionRecord,
      }),
    });

    try {
      const response = await app.inject({
        method: "POST",
        url: "/v2/documents/doc-1/chat/threads/thr-2/runs/run-1/replay",
      });

      expect(response.statusCode).toBe(200);
      expect(createSession).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          body: {
            threadId: "thr-2",
          },
        }),
      );
      expect(replayExecutionRecord).toHaveBeenCalledWith(
        expect.objectContaining({
          documentId: "doc-1",
          threadId: "thr-2",
          recordId: "run-1",
        }),
      );
      expect(response.json()).toEqual(
        expect.objectContaining({
          reviewQueueItems: [expect.objectContaining({ id: "bundle-replay-1" })],
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("returns a structured not-found envelope when the chat thread event stream is stale", async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        async createSession() {
          throw createWorkbookAgentServiceError({
            code: "WORKBOOK_AGENT_SESSION_NOT_FOUND",
            message: "Workbook agent session not found",
            statusCode: 404,
            retryable: true,
          });
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/documents/doc-1/chat/threads/thr-1/events",
      });

      expect(response.statusCode).toBe(404);
      expect(response.json()).toEqual(
        expect.objectContaining({
          error: "WORKBOOK_AGENT_SESSION_NOT_FOUND",
          message: "Workbook agent session not found",
          retryable: true,
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("includes workbook agent observability in healthz when the service is enabled", async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        getObservabilitySnapshot() {
          return {
            enabled: true,
            generatedAtUnixMs: 42,
            featureFlags: {
              sharedThreadsEnabled: true,
              workflowRunnerEnabled: true,
              autoApplyLowRiskEnabled: false,
              formulaWorkflowFamilyEnabled: true,
              formattingWorkflowFamilyEnabled: true,
              importWorkflowFamilyEnabled: true,
              rollupWorkflowFamilyEnabled: true,
              structuralWorkflowFamilyEnabled: true,
              allowlistedUserCount: 2,
              allowlistedDocumentCount: 1,
            },
            sessions: {
              sessionCount: 3,
              subscriberThreadCount: 2,
              subscriberCount: 4,
              activeTurnCount: 1,
              runningWorkflowCount: 1,
              pendingBundleCount: 1,
              sharedPendingReviewCount: 1,
            },
            pool: {
              slotCount: 1,
              boundThreadCount: 2,
              activeTurnCount: 1,
              queuedTurnCount: 0,
              maxClients: 4,
              maxConcurrentTurnsPerClient: 1,
              maxQueuedTurnsPerClient: 8,
            },
            counters: {
              turnBackpressureCount: 1,
              workflowStartedCount: 2,
              workflowCompletedCount: 1,
              workflowFailedCount: 0,
              workflowCancelledCount: 0,
              sharedReviewApprovedCount: 0,
              sharedReviewRejectedCount: 0,
              sharedRecommendationApprovedCount: 1,
              sharedRecommendationRejectedCount: 0,
            },
          };
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/healthz",
      });

      expect(response.statusCode).toBe(200);
      expect(response.json()).toEqual(
        expect.objectContaining({
          ok: true,
          workbookAgent: expect.objectContaining({
            enabled: true,
            generatedAtUnixMs: 42,
            featureFlags: expect.objectContaining({
              allowlistedUserCount: 2,
              allowlistedDocumentCount: 1,
            }),
            sessions: expect.objectContaining({
              sessionCount: 3,
              sharedPendingReviewCount: 1,
            }),
          }),
        }),
      );
    } finally {
      await app.close();
    }
  });

  it("exposes the workbook agent observability snapshot route", async () => {
    const { app } = createSyncServer({
      logger: false,
      workbookAgentService: createWorkbookAgentServiceStub({
        getObservabilitySnapshot() {
          return {
            enabled: true,
            generatedAtUnixMs: 99,
            featureFlags: {
              sharedThreadsEnabled: true,
              workflowRunnerEnabled: true,
              autoApplyLowRiskEnabled: true,
              formulaWorkflowFamilyEnabled: true,
              formattingWorkflowFamilyEnabled: true,
              importWorkflowFamilyEnabled: true,
              rollupWorkflowFamilyEnabled: true,
              structuralWorkflowFamilyEnabled: true,
              allowlistedUserCount: 0,
              allowlistedDocumentCount: 0,
            },
            sessions: {
              sessionCount: 0,
              subscriberThreadCount: 0,
              subscriberCount: 0,
              activeTurnCount: 0,
              runningWorkflowCount: 0,
              pendingBundleCount: 0,
              sharedPendingReviewCount: 0,
            },
            pool: {
              slotCount: 0,
              boundThreadCount: 0,
              activeTurnCount: 0,
              queuedTurnCount: 0,
              maxClients: 4,
              maxConcurrentTurnsPerClient: 1,
              maxQueuedTurnsPerClient: 8,
            },
            counters: {
              turnBackpressureCount: 0,
              workflowStartedCount: 0,
              workflowCompletedCount: 0,
              workflowFailedCount: 0,
              workflowCancelledCount: 0,
              sharedReviewApprovedCount: 0,
              sharedReviewRejectedCount: 0,
              sharedRecommendationApprovedCount: 0,
              sharedRecommendationRejectedCount: 0,
            },
          };
        },
      }),
    });

    try {
      const response = await app.inject({
        method: "GET",
        url: "/v2/agent/observability",
        headers: {
          cookie: "bilig_session=test",
        },
      });

      expect(response.statusCode).toBe(200);
      expect(response.headers["cache-control"]).toBe("no-store");
      expect(response.json()).toEqual(
        expect.objectContaining({
          enabled: true,
          generatedAtUnixMs: 99,
          pool: expect.objectContaining({
            maxClients: 4,
          }),
        }),
      );
    } finally {
      await app.close();
    }
  });
});
