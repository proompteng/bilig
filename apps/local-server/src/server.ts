import Fastify, { type FastifyReply, type FastifyRequest } from "fastify";
import websocket from "@fastify/websocket";

import { decodeAgentFrame, encodeAgentFrame } from "@bilig/agent-api";
import { decodeFrame, encodeFrame } from "@bilig/binary-protocol";

import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";

function noop(): void {}

export interface LocalServerOptions {
  sessionManager?: LocalWorkbookSessionManager;
  logger?: boolean;
}

export function createLocalServer(options: LocalServerOptions = {}) {
  const sessionManager = options.sessionManager ?? new LocalWorkbookSessionManager();
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
    service: "bilig-local-server",
  }));

  app.get(
    "/v1/documents/:documentId/state",
    async (request: FastifyRequest<{ Params: { documentId: string } }>) => {
      return sessionManager.getDocumentState(request.params.documentId);
    },
  );

  app.post(
    "/v1/agent/frames",
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
        : await sessionManager.handleAgentFrame(frame);
      reply.header("content-type", "application/octet-stream");
      return Buffer.from(encodeAgentFrame(response));
    },
  );

  app.register(async (wsApp) => {
    wsApp.get("/v1/documents/:documentId/ws", { websocket: true }, (socket) => {
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
            detach = sessionManager.attachBrowser(documentId, subscriberId, (nextFrame) => {
              ws.send(encodeFrame(nextFrame));
            });
          }
          const responses = await sessionManager.handleSyncFrame(frame);
          responses.forEach((responseFrame) => ws.send(encodeFrame(responseFrame)));
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

type NormalizedWebSocket = {
  on(event: string, listener: (...args: unknown[]) => void): void;
  send(data: Uint8Array): void;
};

function isNormalizedWebSocket(value: unknown): value is NormalizedWebSocket {
  return (
    typeof value === "object" &&
    value !== null &&
    "on" in value &&
    typeof value.on === "function" &&
    "send" in value &&
    typeof value.send === "function"
  );
}

function hasSocket(value: unknown): value is { socket: unknown } {
  return typeof value === "object" && value !== null && "socket" in value;
}

function hasWebSocket(value: unknown): value is { websocket: unknown } {
  return typeof value === "object" && value !== null && "websocket" in value;
}

type EventTargetWebSocket = {
  addEventListener(event: string, listener: (event: unknown) => void): void;
  send(data: Uint8Array): void;
};

function isEventTargetWebSocket(value: unknown): value is EventTargetWebSocket {
  return (
    typeof value === "object" &&
    value !== null &&
    "addEventListener" in value &&
    typeof value.addEventListener === "function" &&
    "send" in value &&
    typeof value.send === "function"
  );
}

function asNormalizedEventTargetSocket(socket: EventTargetWebSocket): NormalizedWebSocket {
  return {
    on(event, listener) {
      socket.addEventListener(event, (payload) => {
        if (
          event === "message" &&
          typeof payload === "object" &&
          payload !== null &&
          "data" in payload
        ) {
          listener(payload.data);
          return;
        }
        listener(payload);
      });
    },
    send(data) {
      socket.send(data);
    },
  };
}

function normalizeWebSocket(candidate: unknown): NormalizedWebSocket {
  if (isNormalizedWebSocket(candidate)) {
    return candidate;
  }
  if (isEventTargetWebSocket(candidate)) {
    return asNormalizedEventTargetSocket(candidate);
  }
  if (hasSocket(candidate) && isNormalizedWebSocket(candidate.socket)) {
    return candidate.socket;
  }
  if (hasSocket(candidate) && isEventTargetWebSocket(candidate.socket)) {
    return asNormalizedEventTargetSocket(candidate.socket);
  }
  if (hasWebSocket(candidate)) {
    return normalizeWebSocket(candidate.websocket);
  }
  throw new Error("Unsupported websocket connection shape");
}

function isStreamingAgentRequest(
  frame: ReturnType<typeof decodeAgentFrame>,
): frame is Extract<ReturnType<typeof decodeAgentFrame>, { kind: "request" }> {
  return (
    frame.kind === "request" &&
    (frame.request.kind === "subscribeRange" || frame.request.kind === "unsubscribe")
  );
}
