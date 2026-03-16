import Fastify from "fastify";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { DocumentSessionManager } from "./document-session-manager.js";

export interface SyncServerOptions {
  sessionManager?: DocumentSessionManager;
}

export function createSyncServer(options: SyncServerOptions = {}) {
  const sessionManager = options.sessionManager ?? new DocumentSessionManager();
  const app = Fastify({ logger: true });

  app.addContentTypeParser("application/octet-stream", { parseAs: "buffer" }, (_request, body, done) => {
    done(null, body);
  });

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-sync-server"
  }));

  app.get("/v1/documents/:documentId/state", async (request) => {
    const params = request.params as { documentId: string };
    return sessionManager.getDocumentState(params.documentId);
  });

  app.get("/v1/documents/:documentId/snapshot/latest", async (request, reply) => {
    const params = request.params as { documentId: string };
    const snapshot = await sessionManager.persistence.snapshots.latest(params.documentId);
    if (!snapshot) {
      reply.code(404);
      return {
        error: "SNAPSHOT_NOT_FOUND"
      };
    }

    reply.header("content-type", snapshot.contentType);
    return Buffer.from(snapshot.bytes);
  });

  app.post("/v1/frames", async (request, reply) => {
    const buffer = request.body as Buffer;
    const response = await sessionManager.handleSyncFrame(decodeFrame(buffer));
    reply.header("content-type", "application/octet-stream");
    return Buffer.from(encodeFrame(response));
  });

  app.post("/v1/agent/frames", async (request, reply) => {
    const buffer = request.body as Buffer;
    const response = await sessionManager.handleAgentFrame(decodeAgentFrame(buffer));
    reply.header("content-type", "application/octet-stream");
    return Buffer.from(encodeAgentFrame(response));
  });

  return { app, sessionManager };
}
