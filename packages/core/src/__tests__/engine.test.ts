import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../index.js";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { EngineOpBatch } from "@bilig/crdt";

describe("SpreadsheetEngine", () => {
  it("recalculates simple formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 20 });

    engine.setCellValue("Sheet1", "A1", 12);
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 24 });
  });

  it("supports cross-sheet references", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.createSheet("Sheet2");
    engine.setCellValue("Sheet1", "A1", 4);
    engine.setCellFormula("Sheet2", "B2", "Sheet1!A1*3");
    expect(engine.getCellValue("Sheet2", "B2")).toEqual({ tag: ValueTag.Number, value: 12 });
  });

  it("uses the wasm fast path for supported aggregate formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 3);
    engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)+ROUND(A1/2)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);
  });

  it("rebinds formulas when a referenced sheet appears later", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellFormula("Sheet1", "A1", "Sheet2!B1*2");

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });

    engine.createSheet("Sheet2");
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 0 });

    engine.setCellValue("Sheet2", "B1", 3);
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 6 });
  });

  it("converges under reordered replicated batches and restores replica state", async () => {
    const engineA = new SpreadsheetEngine({ workbookName: "spec", replicaId: "a" });
    const engineB = new SpreadsheetEngine({ workbookName: "spec", replicaId: "b" });
    await Promise.all([engineA.ready(), engineB.ready()]);

    const outboundA: EngineOpBatch[] = [];
    const outboundB: EngineOpBatch[] = [];
    engineA.subscribeBatches((batch) => outboundA.push(batch));
    engineB.subscribeBatches((batch) => outboundB.push(batch));

    engineA.createSheet("Sheet1");
    engineB.createSheet("Sheet1");
    engineA.setCellValue("Sheet1", "A1", 1);
    engineB.setCellValue("Sheet1", "A1", 2);

    [...outboundB].reverse().forEach((batch) => engineA.applyRemoteBatch(batch));
    [...outboundA].forEach((batch) => engineB.applyRemoteBatch(batch));

    expect(engineA.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engineB.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 2 });

    const restored = new SpreadsheetEngine({ workbookName: "restored", replicaId: "b" });
    await restored.ready();
    restored.importSnapshot(engineB.exportSnapshot());
    restored.importReplicaSnapshot(engineB.exportReplicaSnapshot());

    restored.applyRemoteBatch(outboundA[outboundA.length - 1]!);
    expect(restored.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 2 });
  });

  it("exports sparse high-row cells without truncating the sheet", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A10002", 7);

    const snapshot = engine.exportSnapshot();
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: "A10002", value: 7 });
  });
});
