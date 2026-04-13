import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
} from "./engine-advanced-metadata-fuzz-helpers.js";

describe("engine data validation fuzz", () => {
  it("preserves data validation metadata across structural edits and snapshot restore", async () => {
    await runProperty({
      suite: "core/data-validations/structural-roundtrip",
      arbitrary: fc.array(metadataStructuralActionArbitrary(["Sheet1"]), {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "validation-fuzz" });
        await engine.ready();
        engine.createSheet("Sheet1");
        engine.setDataValidation({
          range: {
            sheetName: "Sheet1",
            startAddress: "B2",
            endAddress: "B4",
          },
          rule: {
            kind: "list",
            values: ["Draft", "Final"],
          },
          allowBlank: false,
        });

        actions.forEach((action) => {
          applyMetadataStructuralAction(engine, action);
        });

        const restored = await restoreMetadataSnapshot(engine, "validation-fuzz-restored");
        expect(restored.getDataValidations("Sheet1")).toEqual(engine.getDataValidations("Sheet1"));
      },
    });
  });
});
