import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import websocket from "@fastify/websocket";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";
import type { RuntimeSession } from "@bilig/contracts";
import {
  createErrorEnvelope,
  createRuntimeSession,
  type DocumentControlService,
  normalizeWebSocket,
  resolveRequestBaseUrl,
  resolveServerRuntimeConfig,
  runPromise,
  toMessageBytes,
} from "@bilig/runtime-kernel";

import { DocumentSessionManager } from "./document-session-manager.js";
import { SyncDocumentSupervisor } from "./document-supervisor.js";
import { resolveRequestSession, resolveSessionIdentity } from "./session.js";
import type { WorksheetExecutor } from "./worksheet-executor.js";
import type { ZeroSyncService } from "./zero/service.js";

export interface SyncServerOptions {
  sessionManager?: DocumentSessionManager;
  documentService?: DocumentControlService;
  worksheetExecutor?: WorksheetExecutor | null;
  zeroSyncService?: ZeroSyncService;
  logger?: boolean;
}

function noop(): void {}

export function createSyncServer(options: SyncServerOptions = {}) {
  const runtimeConfig = resolveServerRuntimeConfig(process.env);
  const sessionManager =
    options.sessionManager ??
    new DocumentSessionManager(undefined, undefined, options.worksheetExecutor ?? null);
  const documentService = options.documentService ?? new SyncDocumentSupervisor(sessionManager);
  const zeroSyncService = options.zeroSyncService;
  const app = Fastify({ logger: options.logger ?? true });
  app.register(websocket);

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer" },
    (_request, body, done) => {
      done(null, body);
    },
  );

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-sync-server",
    zeroSync: zeroSyncService?.enabled ?? false,
  }));

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

  app.register(async (wsApp) => {
    wsApp.get("/v2/documents/:documentId/ws", { websocket: true }, (socket) => {
      const ws = normalizeWebSocket(socket);
      let documentId: string | null = null;
      let sessionId: string | null = null;
      const subscriberId = `sync-browser:${Date.now()}:${Math.random().toString(36).slice(2)}`;
      let detach = noop;

      ws.on("message", async (raw: unknown) => {
        try {
          const frame = decodeFrame(toMessageBytes(raw));
          if (frame.kind === "hello" && documentId === null) {
            documentId = frame.documentId;
            sessionId = frame.sessionId;
            detach = await runPromise(
              documentService.attachBrowser(documentId, subscriberId, (nextFrame) => {
                ws.send(encodeFrame(nextFrame));
              }),
            );
            const helloFrames = await runPromise(documentService.openBrowserSession(frame));
            helloFrames.forEach((responseFrame) => ws.send(encodeFrame(responseFrame)));
            return;
          }
          const response = await runPromise(documentService.handleSyncFrame(frame));
          (Array.isArray(response) ? response : [response]).forEach((responseFrame) =>
            ws.send(encodeFrame(responseFrame)),
          );
        } catch (error) {
          ws.send(
            encodeFrame({
              kind: "error",
              documentId: documentId ?? "unknown",
              code: "SYNC_SERVER_MESSAGE_FAILURE",
              message: error instanceof Error ? error.message : String(error),
              retryable: false,
            }),
          );
        }
      });

      ws.on("close", () => {
        detach();
        if (documentId && sessionId) {
          void sessionManager.persistence.presence.leave(documentId, sessionId);
        }
      });
    });
  });

  return { app, sessionManager, documentService };
}
