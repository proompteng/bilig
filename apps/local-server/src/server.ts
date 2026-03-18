import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import websocket from "@fastify/websocket";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";

export interface LocalServerOptions {
  sessionManager?: LocalWorkbookSessionManager;
}

export function createLocalServer(options: LocalServerOptions = {}) {
  const sessionManager = options.sessionManager ?? new LocalWorkbookSessionManager();
  const app = Fastify({ logger: true });

  app.register(websocket);

  app.addContentTypeParser("application/octet-stream", { parseAs: "buffer" }, (_request, body, done) => {
    done(null, body);
  });

  app.get("/healthz", async () => ({
    ok: true,
    service: "bilig-local-server"
  }));

  app.get("/v1/documents/:documentId/state", async (request: FastifyRequest<{ Params: { documentId: string } }>) => {
    const params = request.params as { documentId: string };
    return sessionManager.getDocumentState(params.documentId);
  });

  app.post("/v1/agent/frames", async (request: FastifyRequest, reply: FastifyReply) => {
    const response = await sessionManager.handleAgentFrame(decodeAgentFrame(request.body as Buffer));
    reply.header("content-type", "application/octet-stream");
    return Buffer.from(encodeAgentFrame(response));
  });

  app.get("/v1/documents/:documentId/ws", { websocket: true }, (socket, request) => {
    const ws = normalizeWebSocket(socket);
    let documentId: string | null = null;
    const subscriberId = `browser:${Date.now()}:${Math.random().toString(36).slice(2)}`;
    let detach = () => {};

    ws.on("message", async (raw: unknown) => {
      try {
        const message = toMessageBytes(raw);
        const frame = decodeFrame(message);
        if (frame.kind === "hello" && documentId === null) {
          documentId = frame.documentId;
          detach = sessionManager.attachBrowser(documentId, subscriberId, (nextFrame) => {
            ws.send(encodeFrame(nextFrame));
          });
        }
        const responses = await sessionManager.handleSyncFrame(frame);
        responses.forEach((frame) => ws.send(encodeFrame(frame)));
      } catch (error) {
        ws.send(
          encodeFrame({
            kind: "error",
            documentId: documentId ?? "unknown",
            code: "LOCAL_SERVER_MESSAGE_FAILURE",
            message: error instanceof Error ? error.message : String(error),
            retryable: false
          })
        );
      }
    });

    ws.on("close", () => {
      detach();
    });
  });

  return { app, sessionManager };
}

function toMessageBytes(raw: unknown): Uint8Array {
  if (raw instanceof Buffer) {
    return new Uint8Array(raw);
  }
  if (raw instanceof ArrayBuffer) {
    return new Uint8Array(raw);
  }
  if (ArrayBuffer.isView(raw)) {
    return new Uint8Array(raw.buffer, raw.byteOffset, raw.byteLength);
  }
  throw new Error("Unsupported websocket payload");
}

function normalizeWebSocket(candidate: unknown): { on: (event: string, listener: (...args: unknown[]) => void) => void; send: (data: Uint8Array) => void } {
  if (typeof candidate === "object" && candidate !== null && "on" in candidate && "send" in candidate) {
    return candidate as { on: (event: string, listener: (...args: unknown[]) => void) => void; send: (data: Uint8Array) => void };
  }
  if (typeof candidate === "object" && candidate !== null && "socket" in candidate) {
    const socket = (candidate as { socket: unknown }).socket;
    if (typeof socket === "object" && socket !== null && "on" in socket && "send" in socket) {
      return socket as { on: (event: string, listener: (...args: unknown[]) => void) => void; send: (data: Uint8Array) => void };
    }
  }
  throw new Error("Unsupported websocket connection shape");
}
