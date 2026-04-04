import { Effect } from "effect";

import type { AgentFrame } from "@bilig/agent-api";
import type { ProtocolFrame } from "@bilig/binary-protocol";
import { TransportError } from "@bilig/runtime-kernel";
import { documentIdFromSessionId } from "./workbook-session-shared.js";

interface ActorLike {
  send(event: unknown): void;
}

export function updateActorFromFrames(
  actor: ActorLike,
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

export function resolveAgentDocumentId(frame: AgentFrame): string | null {
  if (frame.kind !== "request") {
    return null;
  }
  if ("documentId" in frame.request && typeof frame.request.documentId === "string") {
    return frame.request.documentId;
  }
  if ("sessionId" in frame.request && typeof frame.request.sessionId === "string") {
    return documentIdFromSessionId(frame.request.sessionId);
  }
  return null;
}

export function wrapTransportPromise<Success>(
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

export function wrapTransportSync<Success>(
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
