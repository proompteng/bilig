import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";
import {
  applyMetadataStructuralAction,
  metadataStructuralActionArbitrary,
  restoreMetadataSnapshot,
} from "./engine-advanced-metadata-fuzz-helpers.js";

describe("engine chart fuzz", () => {
  it("preserves chart metadata semantics across structural edits and snapshot restore", async () => {
    await runProperty({
      suite: "core/charts/structural-roundtrip",
      arbitrary: fc.array(metadataStructuralActionArbitrary(["Data", "Dashboard"]), {
        minLength: 1,
        maxLength: 10,
      }),
      predicate: async (actions) => {
        const engine = new SpreadsheetEngine({ workbookName: "charts-fuzz" });
        await engine.ready();
        engine.createSheet("Data");
        engine.createSheet("Dashboard");
        engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
          ["Month", "Revenue"],
          ["Jan", 10],
          ["Feb", 15],
          ["Mar", 9],
        ]);
        engine.setChart({
          id: "Trend",
          sheetName: "Dashboard",
          address: "B2",
          source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
          chartType: "line",
          rows: 10,
          cols: 6,
        });

        actions.forEach((action) => {
          applyMetadataStructuralAction(engine, action);
        });

        const restored = await restoreMetadataSnapshot(engine, "charts-fuzz-restored");
        expect(restored.getCharts()).toEqual(engine.getCharts());
        expect(restored.exportSnapshot().workbook.metadata?.charts).toEqual(
          engine.exportSnapshot().workbook.metadata?.charts,
        );
      },
    });
  });
});
