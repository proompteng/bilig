import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
} from "./engine-advanced-metadata-fuzz-helpers.js";

describe("engine media fuzz", () => {
  it("preserves image and shape metadata across structural edits and snapshot restore", async () => {
    await runProperty({
      suite: "core/media/structural-roundtrip",
      arbitrary: fc.array(metadataStructuralActionArbitrary(["Dashboard"]), {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "media-fuzz" });
        await engine.ready();
        engine.createSheet("Dashboard");
        engine.setImage({
          id: "Revenue Image",
          sheetName: "Dashboard",
          address: "B2",
          sourceUrl: "https://example.com/revenue.png",
          rows: 9,
          cols: 6,
        });
        engine.setShape({
          id: "Review Callout",
          sheetName: "Dashboard",
          address: "E5",
          shapeType: "roundedRectangle",
          rows: 3,
          cols: 4,
          text: "Review",
        });

        actions.forEach((action) => {
          applyMetadataStructuralAction(engine, action);
        });

        const restored = await restoreMetadataSnapshot(engine, "media-fuzz-restored");
        expect(restored.getImages()).toEqual(engine.getImages());
        expect(restored.getShapes()).toEqual(engine.getShapes());
      },
    });
  });
});
