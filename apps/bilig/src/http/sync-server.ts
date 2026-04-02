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
import type { BiligRuntimeConfig } from "@bilig/zero-sync";

import { DocumentSessionManager } from "../workbook-runtime/document-session-manager.js";
import { SyncDocumentSupervisor } from "../workbook-runtime/sync-document-supervisor.js";
import { resolveRequestSession, resolveSessionIdentity } from "./session.js";
import type { WorksheetExecutor } from "../workbook-runtime/worksheet-executor.js";
import type { ZeroSyncService } from "../zero/service.js";

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
  logger?: boolean;
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

function resolveWebRuntimeConfig(env: Record<string, string | undefined>): BiligRuntimeConfig {
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
  const app = Fastify({ logger: options.logger ?? true });

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer" },
    (_request, body, done) => {
      done(null, body);
    },
  );

  if (zeroProxyUpstream) {
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
    reply.header("cache-control", "no-store");
    return webRuntimeConfig;
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
