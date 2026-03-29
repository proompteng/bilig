import { describe, expect, it } from "vitest";
import fc from "fast-check";
import type { ProtocolFrame } from "@bilig/binary-protocol";
import { DocumentSessionManager } from "../document-session-manager.js";
import { runScheduledProperty } from "@bilig/test-fuzz";

type SyncAction =
  | { kind: "hello"; sessionId: string }
  | { kind: "appendBatch"; batchId: string }
  | { kind: "heartbeat" }
  | { kind: "inspect" };

function toFrame(documentId: string, action: SyncAction): ProtocolFrame {
  switch (action.kind) {
    case "hello":
      return {
        kind: "hello",
        documentId,
        replicaId: `browser:${action.sessionId}`,
        sessionId: action.sessionId,
        protocolVersion: 1,
        lastServerCursor: 0,
        capabilities: ["browser-sync"],
      };
    case "appendBatch":
      return {
        kind: "appendBatch",
        documentId,
        cursor: 0,
        batch: {
          id: action.batchId,
          replicaId: "replica:fuzz",
          clock: { counter: Number(action.batchId.split("-")[1] ?? "0") + 1 },
          ops: [{ kind: "upsertWorkbook", name: "fuzz-doc" }],
        },
      };
    case "heartbeat":
      return {
        kind: "heartbeat",
        documentId,
        cursor: 0,
        sentAtUnixMs: Date.now(),
      };
    case "inspect":
      return {
        kind: "heartbeat",
        documentId,
        cursor: 0,
        sentAtUnixMs: Date.now(),
      };
  }
}

describe("document session manager fuzz", () => {
  it("keeps cursor and session state coherent across scheduled interleavings", async () => {
    await runScheduledProperty({
      suite: "sync-server/document-session-manager/interleavings",
      arbitrary: fc.array(
        fc.oneof<SyncAction>(
          fc.constantFrom(
            { kind: "hello", sessionId: "session-a" },
            { kind: "hello", sessionId: "session-b" },
            { kind: "heartbeat" },
            { kind: "inspect" },
          ),
          fc.integer({ min: 0, max: 5 }).map((index) => ({
            kind: "appendBatch" as const,
            batchId: `batch-${index}`,
          })),
        ),
        { minLength: 3, maxLength: 10 },
      ),
      predicate: async ({ scheduler, value: actions }) => {
        const manager = new DocumentSessionManager();
        const documentId = `sync-fuzz-${actions.length}`;
        const scheduledHandleFrame = scheduler.scheduleFunction(
          async (frame: ProtocolFrame) => await manager.handleSyncFrame(frame),
        );
        const scheduledInspect = scheduler.scheduleFunction(
          async () => await manager.getDocumentState(documentId),
        );

        const resultPromises = actions.map((action) =>
          action.kind === "inspect"
            ? scheduledInspect()
            : scheduledHandleFrame(toFrame(documentId, action)),
        );
        await scheduler.waitAll();
        const results = await Promise.all(resultPromises);

        const state = await manager.getDocumentState(documentId);
        const acceptedAppendBatches = actions.filter(
          (action) => action.kind === "appendBatch",
        ).length;
        expect(state.cursor).toBe(acceptedAppendBatches);
        expect(new Set(state.sessions).size).toBe(state.sessions.length);
        results.forEach((result) => {
          expect(result.documentId).toBe(documentId);
        });
      },
    });
  }, 20_000);
});
