import { createActor, type ActorRefFrom } from "xstate";
import { Effect } from "effect";

import { createDocumentSupervisorMachine } from "@bilig/actors";
import type { AgentFrame } from "@bilig/agent-api";
import type { HelloFrame, ProtocolFrame } from "@bilig/binary-protocol";
import type { DocumentStateSummary } from "@bilig/contracts";
import {
  type AgentFrameContext,
  type DocumentControlService,
  type SnapshotPayload,
  TransportError,
} from "@bilig/runtime-kernel";

import { LocalWorkbookSessionManager } from "./local-workbook-session-manager.js";

type DocumentSupervisorActor = ActorRefFrom<ReturnType<typeof createDocumentSupervisorMachine>>;

export class LocalDocumentSupervisor implements DocumentControlService {
  private readonly actors = new Map<string, DocumentSupervisorActor>();

  constructor(private readonly manager: LocalWorkbookSessionManager) {}

  attachBrowser(
    documentId: string,
    subscriberId: string,
    send: (frame: ProtocolFrame) => void,
  ): Effect.Effect<() => void, TransportError> {
    return wrapSync("Failed to attach browser subscriber", () => {
      const actor = this.ensureActor(documentId);
      actor.send({ type: "browser.attached" });
      const detach = this.manager.attachBrowser(documentId, subscriberId, send);
      return () => {
        detach();
        actor.send({ type: "browser.detached" });
      };
    });
  }

  openBrowserSession(frame: HelloFrame): Effect.Effect<ProtocolFrame[], TransportError> {
    return this.handleSyncFrame(frame).pipe(
      Effect.map((result) => (Array.isArray(result) ? result : [result])),
    );
  }

  handleSyncFrame(
    frame: ProtocolFrame,
  ): Effect.Effect<ProtocolFrame | ProtocolFrame[], TransportError> {
    return wrapPromise("Failed to handle sync frame", async () => {
      const actor = this.ensureActor(frame.documentId);
      const responses = await this.manager.handleSyncFrame(frame);
      actor.send({ type: "operation.recorded", operation: `sync:${frame.kind}` });
      updateFromFrames(actor, responses);
      return responses;
    });
  }

  handleAgentFrame(
    frame: AgentFrame,
    context: AgentFrameContext = {},
  ): Effect.Effect<AgentFrame, TransportError> {
    return wrapPromise("Failed to handle agent frame", async () => {
      const documentId = resolveAgentDocumentId(frame);
      const actor = documentId ? this.ensureActor(documentId) : null;
      const response = await this.manager.handleAgentFrame(frame, context);
      actor?.send({
        type: "operation.recorded",
        operation: frame.kind === "request" ? `agent:${frame.request.kind}` : `agent:${frame.kind}`,
      });
      if (response.kind === "response" && response.response.kind === "workbookLoaded") {
        this.ensureActor(response.response.documentId).send({
          type: "snapshot.updated",
          cursor: 1,
        });
      }
      return response;
    });
  }

  getDocumentState(documentId: string): Effect.Effect<DocumentStateSummary, TransportError> {
    return wrapSync("Failed to get document state", () => {
      const actor = this.ensureActor(documentId);
      const summary = this.manager.getDocumentState(documentId);
      actor.send({ type: "cursor.updated", cursor: summary.cursor });
      const latestSnapshotCursor = this.manager.getLatestSnapshot(documentId)?.cursor ?? null;
      if (latestSnapshotCursor !== null) {
        actor.send({ type: "snapshot.updated", cursor: latestSnapshotCursor });
      }
      return {
        documentId: summary.documentId,
        cursor: summary.cursor,
        owner: null,
        sessions: [...summary.browserSessions, ...summary.agentSessions],
        latestSnapshotCursor,
      } satisfies DocumentStateSummary;
    });
  }

  getLatestSnapshot(documentId: string): Effect.Effect<SnapshotPayload | null, TransportError> {
    return wrapSync("Failed to load latest snapshot", () => {
      const snapshot = this.manager.getLatestSnapshot(documentId);
      if (!snapshot) {
        return null;
      }
      this.ensureActor(documentId).send({
        type: "snapshot.updated",
        cursor: snapshot.cursor,
      });
      return snapshot satisfies SnapshotPayload;
    });
  }

  private ensureActor(documentId: string): DocumentSupervisorActor {
    const existing = this.actors.get(documentId);
    if (existing) {
      return existing;
    }
    const actor = createActor(createDocumentSupervisorMachine(documentId));
    actor.start();
    this.actors.set(documentId, actor);
    return actor;
  }
}

function updateFromFrames(
  actor: DocumentSupervisorActor,
  frames: ProtocolFrame | ProtocolFrame[],
): void {
  const nextFrames = Array.isArray(frames) ? frames : [frames];
  nextFrames.forEach((frame) => {
    if ("cursor" in frame && typeof frame.cursor === "number") {
      actor.send({ type: "cursor.updated", cursor: frame.cursor });
    }
    if (frame.kind === "cursorWatermark") {
      actor.send({ type: "snapshot.updated", cursor: frame.compactedCursor });
    }
    if (frame.kind === "error") {
      actor.send({ type: "error.raised", message: frame.message });
    }
  });
}

function resolveAgentDocumentId(frame: AgentFrame): string | null {
  if (frame.kind !== "request") {
    return null;
  }
  if ("documentId" in frame.request && typeof frame.request.documentId === "string") {
    return frame.request.documentId;
  }
  if ("sessionId" in frame.request && typeof frame.request.sessionId === "string") {
    return frame.request.sessionId.split(":")[0] ?? null;
  }
  return null;
}

function wrapPromise<Success>(
  message: string,
  run: () => Promise<Success>,
): Effect.Effect<Success, TransportError> {
  return Effect.tryPromise({
    try: run,
    catch: (cause) =>
      new TransportError({
        message,
        cause,
      }),
  });
}

function wrapSync<Success>(
  message: string,
  run: () => Success,
): Effect.Effect<Success, TransportError> {
  return Effect.try({
    try: run,
    catch: (cause) =>
      new TransportError({
        message,
        cause,
      }),
  });
}
