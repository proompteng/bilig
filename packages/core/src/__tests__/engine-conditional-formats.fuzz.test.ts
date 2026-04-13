import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
} from "./engine-advanced-metadata-fuzz-helpers.js";

describe("engine conditional format fuzz", () => {
  it("preserves conditional format metadata across structural edits and snapshot restore", async () => {
    await runProperty({
      suite: "core/conditional-formats/structural-roundtrip",
      arbitrary: fc.array(metadataStructuralActionArbitrary(["Sheet1"]), {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "conditional-format-fuzz" });
        await engine.ready();
        engine.createSheet("Sheet1");
        engine.setConditionalFormat({
          id: "cf-1",
          range: {
            sheetName: "Sheet1",
            startAddress: "B2",
            endAddress: "B4",
          },
          rule: {
            kind: "textContains",
            text: "urgent",
          },
          style: {
            font: { bold: true },
          },
        });

        actions.forEach((action) => {
          applyMetadataStructuralAction(engine, action);
        });

        const restored = await restoreMetadataSnapshot(engine, "conditional-format-fuzz-restored");
        expect(restored.getConditionalFormats("Sheet1")).toEqual(
          engine.getConditionalFormats("Sheet1"),
        );
      },
    });
  });
});
