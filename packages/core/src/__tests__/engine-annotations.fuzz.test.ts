import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
} from "./engine-advanced-metadata-fuzz-helpers.js";

describe("engine annotations fuzz", () => {
  it("preserves comment threads and notes across structural edits and snapshot restore", async () => {
    await runProperty({
      suite: "core/annotations/structural-roundtrip",
      arbitrary: fc.array(metadataStructuralActionArbitrary(["Sheet1"]), {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "annotations-fuzz" });
        await engine.ready();
        engine.createSheet("Sheet1");
        engine.setCommentThread({
          threadId: "thread-1",
          sheetName: "Sheet1",
          address: "B2",
          comments: [{ id: "comment-1", body: "Check this total." }],
        });
        engine.setNote({
          sheetName: "Sheet1",
          address: "C3",
          text: "Manual override",
        });

        actions.forEach((action) => {
          applyMetadataStructuralAction(engine, action);
        });

        const restored = await restoreMetadataSnapshot(engine, "annotations-fuzz-restored");
        expect(restored.getCommentThreads("Sheet1")).toEqual(engine.getCommentThreads("Sheet1"));
        expect(restored.getNotes("Sheet1")).toEqual(engine.getNotes("Sheet1"));
      },
    });
  });
});
