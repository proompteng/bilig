import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import websocket from "@fastify/websocket";
import type { RuntimeSession } from "@bilig/contracts";
import {
  createErrorEnvelope,
  createGuestRuntimeSession,
  type DocumentControlService,
  normalizeWebSocket,
  resolveRequestBaseUrl,
  resolveServerRuntimeConfig,
  runPromise,
  toMessageBytes,
} from "@bilig/runtime-kernel";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { LocalDocumentSupervisor } from "./document-supervisor.js";
import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";

function noop(): void {}

function applyCorsHeaders(reply: FastifyReply, allowOrigin: string): void {
  reply.header("access-control-allow-origin", allowOrigin);
  reply.header("access-control-allow-methods", "GET,POST,OPTIONS");
  reply.header("access-control-allow-headers", "content-type");
  reply.header("access-control-expose-headers", "x-bilig-snapshot-cursor");
  if (allowOrigin !== "*") {
    reply.header("vary", "origin");
  }
}

export interface LocalServerOptions {
  sessionManager?: LocalWorkbookSessionManager;
  documentService?: DocumentControlService;
  logger?: boolean;
}

export function createLocalServer(options: LocalServerOptions = {}) {
  const runtimeConfig = resolveServerRuntimeConfig(process.env);
  const allowOrigin = runtimeConfig.corsOrigin ?? "*";
  const sessionManager = options.sessionManager ?? new LocalWorkbookSessionManager();
  const documentService = options.documentService ?? new LocalDocumentSupervisor(sessionManager);
  const app = Fastify({ logger: options.logger ?? true });
  app.register(websocket);

  app.addHook("onRequest", async (request, reply) => {
    applyCorsHeaders(reply, allowOrigin);
    if (request.method === "OPTIONS") {
      return reply.code(204).send();
    }
  });

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer" },
    (_request, body, done) => {
      done(null, body);
    },
  );

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-local-server",
  }));

  app.get("/v2/session", async () => {
    return createGuestRuntimeSession("guest:local-server") satisfies RuntimeSession;
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
    "/v2/agent/frames",
    async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
      const frame = decodeAgentFrame(request.body);
      const response = isStreamingAgentRequest(frame)
        ? {
            kind: "response" as const,
            response: {
              kind: "error" as const,
              id: frame.request.id,
              code: "AGENT_STREAM_REQUIRES_STREAMING_TRANSPORT",
              message: `${frame.request.kind} requires a streaming agent transport such as stdio`,
              retryable: false,
            },
          }
        : await documentService
            .handleAgentFrame(frame, {
              serverUrl: resolveRequestBaseUrl(request, "127.0.0.1:4381"),
              ...(runtimeConfig.browserAppBaseUrl
                ? { browserAppBaseUrl: runtimeConfig.browserAppBaseUrl }
                : {}),
            })
            .pipe(runPromise);
      reply.header("content-type", "application/octet-stream");
      return Buffer.from(encodeAgentFrame(response));
    },
  );

  app.register(async (wsApp) => {
    wsApp.get("/v2/documents/:documentId/ws", { websocket: true }, (socket) => {
      const ws = normalizeWebSocket(socket);
      let documentId: string | null = null;
      const subscriberId = `browser:${Date.now()}:${Math.random().toString(36).slice(2)}`;
      let detach = noop;

      ws.on("message", async (raw: unknown) => {
        try {
          const message = toMessageBytes(raw);
          const frame = decodeFrame(message);
          if (frame.kind === "hello" && documentId === null) {
            documentId = frame.documentId;
            detach = await documentService
              .attachBrowser(documentId, subscriberId, (nextFrame) => {
                ws.send(encodeFrame(nextFrame));
              })
              .pipe(runPromise);
          }
          const responses = await documentService.handleSyncFrame(frame).pipe(runPromise);
          (Array.isArray(responses) ? responses : [responses]).forEach((responseFrame) =>
            ws.send(encodeFrame(responseFrame)),
          );
        } catch (error) {
          ws.send(
            encodeFrame({
              kind: "error",
              documentId: documentId ?? "unknown",
              code: "LOCAL_SERVER_MESSAGE_FAILURE",
              message: error instanceof Error ? error.message : String(error),
              retryable: false,
            }),
          );
        }
      });

      ws.on("close", () => {
        detach();
      });
    });
  });

  return { app, sessionManager, documentService };
}

function isStreamingAgentRequest(
  frame: ReturnType<typeof decodeAgentFrame>,
): frame is Extract<ReturnType<typeof decodeAgentFrame>, { kind: "request" }> {
  return (
    frame.kind === "request" &&
    (frame.request.kind === "subscribeRange" || frame.request.kind === "unsubscribe")
  );
}
