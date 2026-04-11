import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { CellFlags } from "../cell-store.js";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineMutationSupportService } from "../engine/services/mutation-support-service.js";

function isEngineMutationSupportService(value: unknown): value is EngineMutationSupportService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "materializeSpill") === "function" &&
    typeof Reflect.get(value, "clearOwnedSpill") === "function" &&
    typeof Reflect.get(value, "removeSheetRuntime") === "function"
  );
}

function getMutationSupportService(engine: SpreadsheetEngine): EngineMutationSupportService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const support = Reflect.get(runtime, "support");
  if (!isEngineMutationSupportService(support)) {
    throw new TypeError("Expected engine mutation support service");
  }
  return support;
}

describe("EngineMutationSupportService", () => {
  it("tracks changed roots and unions through the public wrapper methods", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "support-wrapper-roots" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    const support = getMutationSupportService(engine);
    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    expect(a1Index).toBeDefined();
    expect(b1Index).toBeDefined();

    Effect.runSync(support.beginMutationCollection());
    const changedInputCount = Effect.runSync(support.markInputChanged(a1Index!, 0));
    const changedFormulaCount = Effect.runSync(support.markFormulaChanged(b1Index!, 0));
    const explicitChangedCount = Effect.runSync(support.markExplicitChanged(a1Index!, 0));

    expect(changedInputCount).toBe(1);
    expect(changedFormulaCount).toBe(1);
    expect(explicitChangedCount).toBe(1);
    expect(Effect.runSync(support.getChangedInputBuffer())[0]).toBe(a1Index);

    const roots = Effect.runSync(
      support.composeMutationRoots(changedInputCount, changedFormulaCount),
    );
    expect(Array.from(roots)).toEqual([a1Index, b1Index]);

    const eventChanges = Effect.runSync(
      support.composeEventChanges(Uint32Array.of(b1Index!), explicitChangedCount),
    );
    expect(Array.from(eventChanges)).toEqual([a1Index, b1Index]);

    const union = Effect.runSync(
      support.unionChangedSets(Uint32Array.of(a1Index!), Uint32Array.of(a1Index!, b1Index!)),
    );
    expect(Array.from(union)).toEqual([a1Index, b1Index]);

    const ordered = Effect.runSync(
      support.composeChangedRootsAndOrdered(
        Uint32Array.of(a1Index!),
        Uint32Array.of(a1Index!, b1Index!),
        2,
      ),
    );
    expect(Array.from(ordered)).toEqual([a1Index, b1Index]);

    Effect.runSync(support.beginMutationCollection());
    expect(Effect.runSync(support.markSpillRootsChanged([a1Index!], 0))).toBe(1);
    Effect.runSync(support.beginMutationCollection());
    expect(Effect.runSync(support.markPivotRootsChanged([a1Index!], 0))).toBe(1);

    const ensuredByName = Effect.runSync(support.ensureCellTracked("Sheet1", "C1"));
    const ensuredByCoords = Effect.runSync(
      support.ensureCellTrackedByCoords(engine.workbook.getSheet("Sheet1")!.id, 0, 2),
    );
    expect(ensuredByCoords).toBe(ensuredByName);

    Effect.runSync(support.resetMaterializedCellScratch(8));
    expect(Effect.runSync(support.syncDynamicRanges(0))).toBe(0);
  });

  it("materializes and clears spill children through the support service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "support-spill" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);

    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    expect(a1Index).toBeDefined();

    const materialized = Effect.runSync(
      getMutationSupportService(engine).materializeSpill(a1Index!, {
        rows: 2,
        cols: 2,
        values: [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    );

    expect(materialized.ownerValue).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 4 });
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([
      { sheetName: "Sheet1", address: "A1", rows: 2, cols: 2 },
    ]);

    Effect.runSync(getMutationSupportService(engine).clearOwnedSpill(a1Index!));

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined();
  });

  it("reports blocked spills and missing sheet removals through the wrappers", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "support-spill-blocked" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "B1", 9);

    const support = getMutationSupportService(engine);
    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    expect(a1Index).toBeDefined();

    const blocked = Effect.runSync(
      support.materializeSpill(a1Index!, {
        rows: 1,
        cols: 2,
        values: [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
        ],
      }),
    );
    expect(blocked.ownerValue).toMatchObject({
      tag: ValueTag.Error,
    });
    expect(blocked.changedCellIndices).toEqual([]);
    expect(Effect.runSync(support.clearOwnedSpill(a1Index!))).toEqual([]);
    expect(Effect.runSync(support.removeSheetRuntime("Missing", 0))).toEqual({
      changedInputCount: 0,
      formulaChangedCount: 0,
      explicitChangedCount: 0,
    });
  });

  it("removes sheet runtime through the service and moves selection to the next sheet", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "support-delete-sheet" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.createSheet("Sheet2");
    engine.setCellValue("Sheet1", "A1", 7);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setSelection("Sheet1", "B2");

    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    expect(a1Index).toBeDefined();

    const removal = Effect.runSync(
      getMutationSupportService(engine).removeSheetRuntime("Sheet1", 0),
    );

    expect(removal.changedInputCount).toBeGreaterThan(0);
    expect(removal.explicitChangedCount).toBeGreaterThan(0);
    expect(engine.workbook.getSheet("Sheet1")).toBeUndefined();
    expect(engine.getSelectionState()).toMatchObject({
      sheetName: "Sheet2",
      address: "A1",
      anchorAddress: "A1",
    });
    expect((engine.workbook.cellStore.flags[a1Index!] & CellFlags.PendingDelete) !== 0).toBe(true);
  });
});
