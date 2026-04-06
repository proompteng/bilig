import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { afterEach, describe, expect, it, vi } from "vitest";
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
      zeroSyncService: {
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
      },
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
      zeroSyncService: {
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
        async loadAuthoritativeEvents() {
          throw new Error("not used");
        },
      },
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
    const createSession = vi.fn(async () => ({
      sessionId: "agent-session-1",
      documentId: "doc-1",
      threadId: "thr-1",
      status: "idle" as const,
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
    }));

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
});
