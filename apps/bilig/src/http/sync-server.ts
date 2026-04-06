import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import fastifyStatic from "@fastify/static";
import httpProxy from "@fastify/http-proxy";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";
import type { RuntimeSession } from "@bilig/contracts";
import {
  createErrorEnvelope,
  createRuntimeSession,
  type DocumentControlService,
  resolveRequestBaseUrl,
  resolveServerRuntimeConfig,
  runPromise,
} from "@bilig/runtime-kernel";
import { workbookScenarioCreateRequestSchema, type BiligRuntimeConfig } from "@bilig/zero-sync";

import { DocumentSessionManager } from "../workbook-runtime/document-session-manager.js";
import { SyncDocumentSupervisor } from "../workbook-runtime/sync-document-supervisor.js";
import { resolveRequestSession, resolveSessionIdentity } from "./session.js";
import type { WorksheetExecutor } from "../workbook-runtime/worksheet-executor.js";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookAgentService } from "../codex-app/workbook-agent-service.js";
import type { WorkbookAgentStreamEvent } from "@bilig/contracts";
import { isWorkbookAgentServiceError } from "../workbook-agent-errors.js";

const SPA_FALLBACK_PREFIXES = [
  "/api/",
  "/v1/",
  "/v2/",
  "/zero",
  "/healthz",
  "/runtime-config.json",
] as const;

export interface SyncServerOptions {
  sessionManager?: DocumentSessionManager;
  documentService?: DocumentControlService;
  worksheetExecutor?: WorksheetExecutor | null;
  zeroSyncService?: ZeroSyncService;
  workbookAgentService?: WorkbookAgentService;
  logger?: boolean;
}

function resolveZeroKeepaliveUrl(upstream: string): URL {
  const url = new URL(upstream);
  url.pathname = "/keepalive";
  url.search = "";
  url.hash = "";
  return url;
}

function resolveBooleanEnv(value: string | undefined, fallback: boolean): boolean {
  const normalized = value?.trim().toLowerCase();
  if (!normalized) {
    return fallback;
  }
  if (normalized === "true" || normalized === "1") {
    return true;
  }
  if (normalized === "false" || normalized === "0") {
    return false;
  }
  throw new Error(`Invalid boolean environment value: ${value}`);
}

function resolveWebRuntimeConfig(
  env: Record<string, string | undefined>,
): Omit<BiligRuntimeConfig, "currentUserId"> {
  const zeroCacheUrl = env["BILIG_ZERO_CACHE_URL"]?.trim() || "/zero";
  const defaultDocumentId = env["BILIG_DEFAULT_DOCUMENT_ID"]?.trim() || "bilig-demo";

  return {
    zeroCacheUrl,
    defaultDocumentId,
    persistState: resolveBooleanEnv(env["BILIG_PERSIST_STATE"], true),
  };
}

function resolveWebDistRoot(): string | null {
  const candidate = join(dirname(fileURLToPath(import.meta.url)), "../../public");
  return existsSync(candidate) ? candidate : null;
}

function shouldServeSpaFallback(method: string, url: string): boolean {
  if (method !== "GET" && method !== "HEAD") {
    return false;
  }

  const pathname = url.split("?", 1)[0] ?? url;
  if (pathname.includes(".", pathname.lastIndexOf("/") + 1)) {
    return false;
  }

  return !SPA_FALLBACK_PREFIXES.some(
    (prefix) => pathname === prefix || pathname.startsWith(prefix),
  );
}

export function createSyncServer(options: SyncServerOptions = {}) {
  const runtimeConfig = resolveServerRuntimeConfig(process.env);
  const webRuntimeConfig = resolveWebRuntimeConfig(process.env);
  const webDistRoot = resolveWebDistRoot();
  const zeroProxyUpstream = process.env["BILIG_ZERO_PROXY_UPSTREAM"]?.trim();
  const sessionManager =
    options.sessionManager ??
    new DocumentSessionManager(undefined, undefined, options.worksheetExecutor ?? null);
  const documentService = options.documentService ?? new SyncDocumentSupervisor(sessionManager);
  const zeroSyncService = options.zeroSyncService;
  const workbookAgentService = options.workbookAgentService;
  const app = Fastify({ logger: options.logger ?? true });
  const handleWorkbookAgentRequest = async <T>(
    request: FastifyRequest,
    reply: FastifyReply,
    task: (
      service: WorkbookAgentService,
      session: ReturnType<typeof resolveSessionIdentity>,
    ) => Promise<T>,
  ): Promise<T | ReturnType<typeof createErrorEnvelope>> => {
    if (!workbookAgentService?.enabled) {
      reply.code(503);
      return createErrorEnvelope(
        "WORKBOOK_AGENT_DISABLED",
        "Workbook agent service is not configured",
        true,
      );
    }
    const session = resolveSessionIdentity(request, reply);
    reply.header("cache-control", "no-store");
    try {
      return await task(workbookAgentService, session);
    } catch (error) {
      if (isWorkbookAgentServiceError(error)) {
        reply.code(error.statusCode);
        return createErrorEnvelope(error.code, error.message, error.retryable);
      }
      throw error;
    }
  };

  app.addHook("onClose", async () => {
    await workbookAgentService?.close().catch(() => undefined);
  });

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer" },
    (_request, body, done) => {
      done(null, body);
    },
  );

  if (zeroProxyUpstream) {
    app.route({
      method: ["GET", "HEAD"],
      url: "/zero/keepalive",
      async handler(request, reply) {
        try {
          const upstreamResponse = await fetch(resolveZeroKeepaliveUrl(zeroProxyUpstream), {
            method: request.method,
            cache: "no-store",
            signal: AbortSignal.timeout(2_000),
          });
          reply.code(upstreamResponse.status);
          reply.header("cache-control", "no-store");
          const contentType = upstreamResponse.headers.get("content-type");
          if (contentType) {
            reply.header("content-type", contentType);
          }
          if (request.method === "HEAD") {
            return reply.send();
          }
          return Buffer.from(await upstreamResponse.arrayBuffer());
        } catch {
          reply.code(503);
          reply.header("cache-control", "no-store");
          return createErrorEnvelope(
            "ZERO_CACHE_UNAVAILABLE",
            "Zero cache keepalive probe failed",
            true,
          );
        }
      },
    });
    app.get("/zero", async (_request, reply) => {
      return reply.redirect("/zero/");
    });
    app.register(httpProxy, {
      upstream: zeroProxyUpstream,
      prefix: "/zero/",
      rewritePrefix: "/",
      websocket: true,
      http2: false,
    });
  }

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-app",
    zeroSync: zeroSyncService?.enabled ?? false,
    web: webDistRoot !== null,
  }));

  app.get("/runtime-config.json", async (_request, reply) => {
    const session = resolveSessionIdentity(_request, reply);
    reply.header("cache-control", "no-store");
    return {
      ...webRuntimeConfig,
      currentUserId: session.userID,
    } satisfies BiligRuntimeConfig;
  });

  app.get(
    "/v2/documents/:documentId/state",
    async (request: FastifyRequest<{ Params: { documentId: string } }>) => {
      return await runPromise(documentService.getDocumentState(request.params.documentId));
    },
  );

  app.get(
    "/v2/documents/:documentId/snapshot/latest",
    async (request: FastifyRequest<{ Params: { documentId: string } }>, reply: FastifyReply) => {
      const snapshot = await runPromise(
        documentService.getLatestSnapshot(request.params.documentId),
      );
      if (!snapshot) {
        reply.code(404);
        return createErrorEnvelope("SNAPSHOT_NOT_FOUND", "Latest snapshot was not found", false);
      }

      reply.header("x-bilig-snapshot-cursor", String(snapshot.cursor));
      reply.header("content-type", snapshot.contentType);
      return Buffer.from(snapshot.bytes);
    },
  );

  app.get(
    "/v2/documents/:documentId/events",
    async (
      request: FastifyRequest<{
        Params: { documentId: string };
        Querystring: { afterRevision?: string };
      }>,
      reply: FastifyReply,
    ) => {
      if (!zeroSyncService?.enabled) {
        reply.code(503);
        return createErrorEnvelope(
          "ZERO_SYNC_DISABLED",
          "Authoritative workbook events require Zero sync",
          true,
        );
      }
      const rawAfterRevision = request.query.afterRevision?.trim() ?? "0";
      const afterRevision = Number.parseInt(rawAfterRevision, 10);
      if (!Number.isFinite(afterRevision) || afterRevision < 0) {
        reply.code(400);
        return createErrorEnvelope(
          "INVALID_AFTER_REVISION",
          "afterRevision must be a non-negative integer",
          false,
        );
      }
      reply.header("cache-control", "no-store");
      return await zeroSyncService.loadAuthoritativeEvents(
        request.params.documentId,
        afterRevision,
      );
    },
  );

  app.post(
    "/v2/documents/:documentId/scenarios",
    async (
      request: FastifyRequest<{
        Params: { documentId: string };
        Body: Record<string, unknown>;
      }>,
      reply: FastifyReply,
    ) => {
      if (!zeroSyncService?.enabled) {
        reply.code(503);
        return createErrorEnvelope("ZERO_SYNC_DISABLED", "Zero sync is not configured", true);
      }
      const session = resolveSessionIdentity(request, reply);
      reply.header("cache-control", "no-store");
      const parsed = workbookScenarioCreateRequestSchema.parse(request.body ?? {});
      return await zeroSyncService.createWorkbookScenario(
        {
          workbookId: request.params.documentId,
          name: parsed.name,
          ...(parsed.sheetName ? { sheetName: parsed.sheetName } : {}),
          ...(parsed.address ? { address: parsed.address } : {}),
          ...(parsed.viewport ? { viewport: parsed.viewport } : {}),
        },
        session,
      );
    },
  );

  app.delete(
    "/v2/documents/:documentId/scenarios/:scenarioDocumentId",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; scenarioDocumentId: string };
      }>,
      reply: FastifyReply,
    ) => {
      if (!zeroSyncService?.enabled) {
        reply.code(503);
        return createErrorEnvelope("ZERO_SYNC_DISABLED", "Zero sync is not configured", true);
      }
      const session = resolveSessionIdentity(request, reply);
      reply.header("cache-control", "no-store");
      await zeroSyncService.deleteWorkbookScenario(
        {
          workbookId: request.params.documentId,
          documentId: request.params.scenarioDocumentId,
        },
        session,
      );
      return { ok: true as const };
    },
  );

  app.post(
    "/v2/documents/:documentId/frames",
    async (
      request: FastifyRequest<{ Params: { documentId: string }; Body: Buffer }>,
      reply: FastifyReply,
    ) => {
      const frame = decodeFrame(request.body);
      if (frame.documentId !== request.params.documentId) {
        reply.code(400);
        return createErrorEnvelope(
          "DOCUMENT_ID_MISMATCH",
          "Frame document id does not match route document id",
          false,
        );
      }
      const response = await runPromise(documentService.handleSyncFrame(frame));
      reply.header("content-type", "application/octet-stream");
      return Buffer.from(encodeFrame(Array.isArray(response) ? (response[0] ?? frame) : response));
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions",
    async (
      request: FastifyRequest<{
        Params: { documentId: string };
        Body: Record<string, unknown>;
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.createSession({
          documentId: request.params.documentId,
          session,
          body: request.body ?? {},
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/context",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string };
        Body: Record<string, unknown>;
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.updateContext({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          session,
          body: request.body ?? {},
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/turns",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string };
        Body: Record<string, unknown>;
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.startTurn({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          session,
          body: request.body ?? {},
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/interrupt",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string };
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.interruptTurn({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          session,
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/bundles/:bundleId/apply",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string; bundleId: string };
        Body: {
          appliedBy?: "user" | "auto";
          commandIndexes?: number[];
          preview?: unknown;
        };
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        const commandIndexes =
          request.body &&
          typeof request.body === "object" &&
          Array.isArray(request.body.commandIndexes)
            ? request.body.commandIndexes
            : undefined;
        return await service.applyPendingBundle({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          bundleId: request.params.bundleId,
          session,
          appliedBy: request.body && request.body.appliedBy === "auto" ? "auto" : "user",
          ...(commandIndexes ? { commandIndexes } : {}),
          preview:
            request.body && typeof request.body === "object" && "preview" in request.body
              ? (request.body.preview ?? null)
              : null,
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/bundles/:bundleId/dismiss",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string; bundleId: string };
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.dismissPendingBundle({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          bundleId: request.params.bundleId,
          session,
        });
      });
    },
  );

  app.post(
    "/v2/documents/:documentId/agent/sessions/:sessionId/runs/:recordId/replay",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string; recordId: string };
      }>,
      reply: FastifyReply,
    ) => {
      return await handleWorkbookAgentRequest(request, reply, async (service, session) => {
        return await service.replayExecutionRecord({
          documentId: request.params.documentId,
          sessionId: request.params.sessionId,
          recordId: request.params.recordId,
          session,
        });
      });
    },
  );

  app.get(
    "/v2/documents/:documentId/agent/sessions/:sessionId/events",
    async (
      request: FastifyRequest<{
        Params: { documentId: string; sessionId: string };
      }>,
      reply: FastifyReply,
    ) => {
      if (!workbookAgentService?.enabled) {
        reply.code(503);
        return createErrorEnvelope(
          "WORKBOOK_AGENT_DISABLED",
          "Workbook agent service is not configured",
          true,
        );
      }
      const session = resolveSessionIdentity(request, reply);
      const sessionSnapshot = workbookAgentService.getSnapshot({
        documentId: request.params.documentId,
        sessionId: request.params.sessionId,
        session,
      });

      reply.hijack();
      const raw = reply.raw;
      raw.writeHead(200, {
        "content-type": "text/event-stream; charset=utf-8",
        "cache-control": "no-store",
        connection: "keep-alive",
      });

      const writeEvent = (event: WorkbookAgentStreamEvent) => {
        raw.write(`data: ${JSON.stringify(event)}\n\n`);
      };

      writeEvent({
        type: "snapshot",
        snapshot: sessionSnapshot,
      });

      const unsubscribe = workbookAgentService.subscribe(request.params.sessionId, (event) => {
        writeEvent(event);
      });
      const keepalive = setInterval(() => {
        raw.write(":keepalive\n\n");
      }, 15_000);

      request.raw.on("close", () => {
        clearInterval(keepalive);
        unsubscribe();
      });
      return reply;
    },
  );

  const handleZeroQuery = async (request: FastifyRequest, reply: FastifyReply) => {
    if (!zeroSyncService?.enabled) {
      reply.code(503);
      return createErrorEnvelope("ZERO_SYNC_DISABLED", "Zero sync is not configured", true);
    }
    resolveSessionIdentity(request, reply);
    return await zeroSyncService.handleQuery(request);
  };

  const handleZeroMutate = async (request: FastifyRequest, reply: FastifyReply) => {
    if (!zeroSyncService?.enabled) {
      reply.code(503);
      return createErrorEnvelope("ZERO_SYNC_DISABLED", "Zero sync is not configured", true);
    }
    resolveSessionIdentity(request, reply);
    return await zeroSyncService.handleMutate(request);
  };

  app.post("/api/zero/v2/query", handleZeroQuery);
  app.post("/api/zero/v2/mutate", handleZeroMutate);

  const handleSessionRequest = async (request: FastifyRequest, reply: FastifyReply) => {
    const session = resolveSessionIdentity(request, reply);
    const requestSession = resolveRequestSession(request);
    return createRuntimeSession({
      authToken: session.userID,
      userId: session.userID,
      roles: requestSession.roles,
      isAuthenticated: requestSession.isAuthenticated,
      authSource: requestSession.authSource,
    }) satisfies RuntimeSession;
  };
  app.get("/v2/session", handleSessionRequest);

  app.post(
    "/v2/agent/frames",
    async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
      const response = await runPromise(
        documentService.handleAgentFrame(decodeAgentFrame(request.body), {
          serverUrl: resolveRequestBaseUrl(request, "127.0.0.1:4321"),
          ...(runtimeConfig.browserAppBaseUrl
            ? { browserAppBaseUrl: runtimeConfig.browserAppBaseUrl }
            : {}),
        }),
      );
      reply.header("content-type", "application/octet-stream");
      return Buffer.from(encodeAgentFrame(response));
    },
  );

  if (webDistRoot) {
    app.register(fastifyStatic, {
      root: webDistRoot,
      prefix: "/",
      maxAge: "30d",
      immutable: true,
    });

    app.get("/", async (_request, reply) => {
      reply.header("cache-control", "no-store");
      return reply.sendFile("index.html", { maxAge: 0, immutable: false });
    });

    app.setNotFoundHandler(async (request, reply) => {
      if (!shouldServeSpaFallback(request.method, request.url)) {
        reply.code(404);
        return createErrorEnvelope("NOT_FOUND", "Route not found", false);
      }

      reply.header("cache-control", "no-store");
      return reply.sendFile("index.html", { maxAge: 0, immutable: false });
    });
  }

  return { app, sessionManager, documentService };
}
