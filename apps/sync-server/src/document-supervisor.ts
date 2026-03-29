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

import { DocumentSessionManager } from "./document-session-manager.js";

type DocumentSupervisorActor = ActorRefFrom<ReturnType<typeof createDocumentSupervisorMachine>>;

export class SyncDocumentSupervisor implements DocumentControlService {
  private readonly actors = new Map<string, DocumentSupervisorActor>();

  constructor(private readonly manager: DocumentSessionManager) {}

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
    return wrapPromise("Failed to open browser session", async () => {
      const actor = this.ensureActor(frame.documentId);
      const responses = await this.manager.openBrowserSession(frame);
      actor.send({ type: "operation.recorded", operation: "openBrowserSession" });
      actor.send({ type: "browser.attached" });
      updateFromFrames(actor, responses);
      return responses;
    });
  }

  handleSyncFrame(
    frame: ProtocolFrame,
  ): Effect.Effect<ProtocolFrame | ProtocolFrame[], TransportError> {
    return wrapPromise("Failed to handle sync frame", async () => {
      const actor = this.ensureActor(frame.documentId);
      const response = await this.manager.handleSyncFrame(frame);
      actor.send({ type: "operation.recorded", operation: `sync:${frame.kind}` });
      updateFromFrames(actor, response);
      return response;
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
      if (
        documentId &&
        response.kind === "response" &&
        response.response.kind === "workbookLoaded"
      ) {
        this.ensureActor(response.response.documentId).send({
          type: "snapshot.updated",
          cursor: 1,
        });
      }
      return response;
    });
  }

  getDocumentState(documentId: string): Effect.Effect<DocumentStateSummary, TransportError> {
    return wrapPromise("Failed to get document state", async () => {
      const actor = this.ensureActor(documentId);
      const state = await this.manager.getDocumentState(documentId);
      actor.send({ type: "cursor.updated", cursor: state.cursor });
      if (state.latestSnapshotCursor !== null) {
        actor.send({ type: "snapshot.updated", cursor: state.latestSnapshotCursor });
      }
      return state;
    });
  }

  getLatestSnapshot(documentId: string): Effect.Effect<SnapshotPayload | null, TransportError> {
    return wrapPromise("Failed to load latest snapshot", async () => {
      const snapshot = await this.manager.persistence.snapshots.latest(documentId);
      if (snapshot) {
        this.ensureActor(documentId).send({
          type: "snapshot.updated",
          cursor: snapshot.cursor,
        });
        return {
          cursor: snapshot.cursor,
          contentType: snapshot.contentType,
          bytes: snapshot.bytes,
        } satisfies SnapshotPayload;
      }
      return null;
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
