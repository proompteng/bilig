import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import { restoreMetadataSnapshot } from "./engine-advanced-metadata-fuzz-helpers.js";

type ProtectionAction =
  | { kind: "setSheet"; hideFormulas: boolean }
  | { kind: "clearSheet" }
  | {
      kind: "setRange";
      id: string;
      startAddress: string;
      endAddress: string;
      hideFormulas: boolean;
    }
  | { kind: "deleteRange"; id: string };

describe("engine protection fuzz", () => {
  it("preserves protection metadata across randomized protection updates and snapshot restore", async () => {
    await runProperty({
      suite: "core/protection/metadata-roundtrip",
      arbitrary: fc.array(protectionActionArbitrary, {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "protection-fuzz" });
        await engine.ready();
        engine.createSheet("Sheet1");

        actions.forEach((action) => {
          switch (action.kind) {
            case "setSheet":
              engine.setSheetProtection({ sheetName: "Sheet1", hideFormulas: action.hideFormulas });
              return;
            case "clearSheet":
              engine.clearSheetProtection("Sheet1");
              return;
            case "setRange":
              engine.setRangeProtection({
                id: action.id,
                range: {
                  sheetName: "Sheet1",
                  startAddress: action.startAddress,
                  endAddress: action.endAddress,
                },
                hideFormulas: action.hideFormulas,
              });
              return;
            case "deleteRange":
              engine.deleteRangeProtection(action.id);
              return;
          }
        });

        const restored = await restoreMetadataSnapshot(engine, "protection-fuzz-restored");
        expect(restored.getSheetProtection("Sheet1")).toEqual(engine.getSheetProtection("Sheet1"));
        expect(restored.getRangeProtections("Sheet1")).toEqual(
          engine.getRangeProtections("Sheet1"),
        );
      },
    });
  });
});

const protectionActionArbitrary = fc.oneof<ProtectionAction>(
  fc.boolean().map((hideFormulas) => ({ kind: "setSheet", hideFormulas })),
  fc.constant({ kind: "clearSheet" }),
  fc
    .record({
      id: fc.constantFrom("protect-a1", "protect-b2", "protect-c3"),
      startAddress: fc.constantFrom("A1", "B2", "C3"),
      endAddress: fc.constantFrom("B2", "C3", "D4"),
      hideFormulas: fc.boolean(),
    })
    .map((action) => Object.assign({ kind: "setRange" as const }, action)),
  fc.constantFrom<ProtectionAction>(
    { kind: "deleteRange", id: "protect-a1" },
    { kind: "deleteRange", id: "protect-b2" },
    { kind: "deleteRange", id: "protect-c3" },
  ),
);
