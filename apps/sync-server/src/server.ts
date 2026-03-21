import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { DocumentSessionManager } from "./document-session-manager.js";
import type { WorksheetExecutor } from "./worksheet-executor.js";

export interface SyncServerOptions {
  sessionManager?: DocumentSessionManager;
  worksheetExecutor?: WorksheetExecutor | null;
  logger?: boolean;
}

export function createSyncServer(options: SyncServerOptions = {}) {
  const sessionManager = options.sessionManager ?? new DocumentSessionManager(undefined, undefined, options.worksheetExecutor ?? null);
  const app = Fastify({ logger: options.logger ?? true });

  app.addContentTypeParser("application/octet-stream", { parseAs: "buffer" }, (_request, body, done) => {
    done(null, body);
  });

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-sync-server"
  }));

  app.get("/v1/documents/:documentId/state", async (request: FastifyRequest<{ Params: { documentId: string } }>) => {
    return sessionManager.getDocumentState(request.params.documentId);
  });

  app.get(
    "/v1/documents/:documentId/snapshot/latest",
    async (request: FastifyRequest<{ Params: { documentId: string } }>, reply: FastifyReply) => {
    const snapshot = await sessionManager.persistence.snapshots.latest(request.params.documentId);
    if (!snapshot) {
      reply.code(404);
      return {
        error: "SNAPSHOT_NOT_FOUND"
      };
    }

    reply.header("content-type", snapshot.contentType);
    return Buffer.from(snapshot.bytes);
    }
  );

  app.post("/v1/frames", async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
    const response = await sessionManager.handleSyncFrame(decodeFrame(request.body));
    reply.header("content-type", "application/octet-stream");
    return Buffer.from(encodeFrame(response));
  });

  app.post("/v1/agent/frames", async (request: FastifyRequest<{ Body: Buffer }>, reply: FastifyReply) => {
    const response = await sessionManager.handleAgentFrame(decodeAgentFrame(request.body));
    reply.header("content-type", "application/octet-stream");
    return Buffer.from(encodeAgentFrame(response));
  });

  return { app, sessionManager };
}
