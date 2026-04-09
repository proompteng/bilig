import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import type { EnginePivotService } from "../engine/services/pivot-service.js";

function isEnginePivotService(value: unknown): value is EnginePivotService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "materializePivot") === "function" &&
    typeof Reflect.get(value, "resolvePivotData") === "function" &&
    typeof Reflect.get(value, "clearOwnedPivot") === "function" &&
    typeof Reflect.get(value, "clearPivotForCell") === "function"
  );
}

function getPivotService(engine: SpreadsheetEngine): EnginePivotService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null || !("pivot" in runtime)) {
    throw new TypeError("Expected engine runtime to expose a pivot service");
  }
  const pivot = Reflect.get(runtime, "pivot");
  if (!isEnginePivotService(pivot)) {
    throw new TypeError("Expected engine runtime pivot service");
  }
  return pivot;
}

async function buildPivotEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({ workbookName: "spec" });
  await engine.ready();
  engine.createSheet("Data");
  engine.createSheet("Pivot");
  engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "D4" }, [
    ["Region", "Notes", "Product", "Sales"],
    ["East", "priority", "Widget", 10],
    ["West", "priority", "Widget", 7],
    ["East", "priority", "Gizmo", 5],
  ]);
  engine.setPivotTable("Pivot", "B2", {
    name: "SalesByRegion",
    source: { sheetName: "Data", startAddress: "A1", endAddress: "D4" },
    groupBy: ["Region"],
    values: [
      { sourceColumn: "Sales", summarizeBy: "sum" },
      { sourceColumn: "Product", summarizeBy: "count", outputLabel: "Rows" },
    ],
  });
  return engine;
}

describe("EnginePivotService", () => {
  it("clears owned pivot output cells without leaving stale ownership behind", async () => {
    const engine = await buildPivotEngine();
    const service = getPivotService(engine);
    const pivot = engine.getPivotTables()[0];
    if (!pivot) {
      throw new TypeError("Expected pivot table");
    }

    const changed = Effect.runSync(service.clearOwnedPivot(pivot));

    expect(changed.length).toBeGreaterThan(0);
    expect(engine.getCellValue("Pivot", "B2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Pivot", "C3")).toEqual({ tag: ValueTag.Empty });
  });

  it("resolves pivot aggregates through the extracted service boundary", async () => {
    const engine = await buildPivotEngine();
    const service = getPivotService(engine);

    const resolved = Effect.runSync(
      service.resolvePivotData("Pivot", "C3", "Sales", [
        { field: "Region", item: engine.getCellValue("Pivot", "B3") },
      ]),
    );

    expect(resolved).toEqual({ tag: ValueTag.Number, value: 15 });
  });
});
