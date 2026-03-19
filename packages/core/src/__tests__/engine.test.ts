import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../index.js";
import { ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
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

  it("evaluates string formulas and string comparisons on the JS path", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", "hello");
    engine.setCellFormula("Sheet1", "B1", "A1&\" world\"");
    engine.setCellFormula("Sheet1", "C1", "A1=\"HELLO\"");
    engine.setCellFormula("Sheet1", "D1", "\"b\">\"A\"");

    expect(engine.getCellValue("Sheet1", "B1")).toMatchObject({ tag: ValueTag.String, value: "hello world" });
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getCellValue("Sheet1", "D1")).toEqual({ tag: ValueTag.Boolean, value: true });
  });

  it("relocates relative formulas when copying a range", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 5);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    engine.copyRange(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "B2" }
    );

    expect(engine.getCell("Sheet1", "B2").formula).toBe("A2*2");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 10 });
  });

  it("preserves absolute references when copying formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 3);
    engine.setCellValue("Sheet1", "A2", 4);
    engine.setCellFormula("Sheet1", "B1", "$A1+A$1+$A$1");

    engine.copyRange(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
      { sheetName: "Sheet1", startAddress: "C2", endAddress: "C2" }
    );

    expect(engine.getCell("Sheet1", "C2").formula).toBe("$A2+B$1+$A$1");
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 16 });
  });

  it("relocates formulas when filling down", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 4);
    engine.setCellFormula("Sheet1", "B1", "A1*3");

    engine.fillRange(
      { sheetName: "Sheet1", startAddress: "B1", endAddress: "B1" },
      { sheetName: "Sheet1", startAddress: "B2", endAddress: "B3" }
    );

    expect(engine.getCell("Sheet1", "B2").formula).toBe("A2*3");
    expect(engine.getCell("Sheet1", "B3").formula).toBe("A3*3");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 12 });
  });

  it("stores invalid formulas as #VALUE errors instead of throwing", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");

    expect(() => engine.setCellFormula("Sheet1", "A1", "1+")).not.toThrow();
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
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

  it("evaluates row and column aggregate ranges on the JS path", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A3", 5);
    engine.setCellValue("Sheet1", "B3", 7);
    engine.setCellFormula("Sheet1", "C1", "SUM(A:A)");
    engine.setCellFormula("Sheet1", "C2", "SUM(3:3)");

    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(engine.getLastMetrics().jsFormulaCount).toBeGreaterThanOrEqual(1);
  });

  it("uses the wasm fast path for supported aggregate formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 3);
    engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)+ABS(A1/2)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);
  });

  it("deduplicates overlapped precedents when scalar refs and ranges touch the same cell", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "A2", 3);
    engine.setCellFormula("Sheet1", "B1", "SUM(A1:A2)+A1");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(engine.getDependencies("Sheet1", "B1").directPrecedents).toEqual(["Sheet1!A1", "Sheet1!A2"]);
  });

  it("keeps branch formulas on the JS path until wasm semantics catch up", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 3);
    engine.setCellValue("Sheet1", "A2", 9);
    engine.setCellFormula("Sheet1", "B1", "IF(A1>0,A1*2,A2-1)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 6 });
    expect(engine.getLastMetrics().jsFormulaCount).toBe(1);

    engine.setCellValue("Sheet1", "A1", 0);
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(engine.getLastMetrics().jsFormulaCount).toBe(1);
  });

  it("uses the wasm fast path for exact-parity logical builtins", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "A2", 0);
    engine.setCellFormula("Sheet1", "B1", "AND(A1,TRUE)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellFormula("Sheet1", "B2", "OR(A2,FALSE)");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellFormula("Sheet1", "B3", "NOT(A2)");
    expect(engine.getCellValue("Sheet1", "B3")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellValue("Sheet1", "A3", "hello");
    engine.setCellFormula("Sheet1", "B4", "AND(A3,TRUE)");
    expect(engine.getCellValue("Sheet1", "B4")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);
  });

  it("uses the wasm fast path for exact-parity info and date builtins", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 42);
    engine.setCellValue("Sheet1", "A2", true);
    engine.setCellValue("Sheet1", "A3", "hello");
    engine.setCellValue("Sheet1", "A4", 45351);
    engine.setCellValue("Sheet1", "A5", 45351.75);
    engine.setCellValue("Sheet1", "A6", 60);
    engine.setCellValue("Sheet1", "A7", 45322);
    engine.setCellValue("Sheet1", "A8", 45337);
    engine.setCellValue("Sheet1", "A9", "bad");

    engine.setCellFormula("Sheet1", "B1", "ISBLANK()");
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B2", "ISNUMBER()");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B3", "ISTEXT()");
    expect(engine.getCellValue("Sheet1", "B3")).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B4", "ISBLANK(A1)");
    expect(engine.getCellValue("Sheet1", "B4")).toEqual({ tag: ValueTag.Boolean, value: false });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B5", "ISNUMBER(A1)");
    expect(engine.getCellValue("Sheet1", "B5")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B6", "ISTEXT(A3)");
    expect(engine.getCellValue("Sheet1", "B6")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B7", "DATE(2024,2,29)");
    expect(engine.getCellValue("Sheet1", "B7")).toEqual({ tag: ValueTag.Number, value: 45351 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B8", "YEAR(B7)");
    expect(engine.getCellValue("Sheet1", "B8")).toEqual({ tag: ValueTag.Number, value: 2024 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B9", "MONTH(A5)");
    expect(engine.getCellValue("Sheet1", "B9")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B10", "DAY(A6)");
    expect(engine.getCellValue("Sheet1", "B10")).toEqual({ tag: ValueTag.Number, value: 29 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B11", "EDATE(A7,1.9)");
    expect(engine.getCellValue("Sheet1", "B11")).toEqual({ tag: ValueTag.Number, value: 45351 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B12", "EOMONTH(A8,A2)");
    expect(engine.getCellValue("Sheet1", "B12")).toEqual({ tag: ValueTag.Number, value: 45382 });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B13", "DATE(A9,2,29)");
    expect(engine.getCellValue("Sheet1", "B13")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B14", "EDATE(A9,1)");
    expect(engine.getCellValue("Sheet1", "B14")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });

    engine.setCellFormula("Sheet1", "B15", "EOMONTH(A9,1)");
    expect(engine.getCellValue("Sheet1", "B15")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 });
  });

  it("uses the wasm fast path for exact-parity rounding builtins", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 123.4);
    engine.setCellFormula("Sheet1", "B1", "ROUND(A1,-1)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 120 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellFormula("Sheet1", "B2", "FLOOR(TRUE,0.5)");
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellFormula("Sheet1", "B3", "CEILING(7,2)");
    expect(engine.getCellValue("Sheet1", "B3")).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellFormula("Sheet1", "B4", "FLOOR(A1,0)");
    expect(engine.getCellValue("Sheet1", "B4")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);

    engine.setCellValue("Sheet1", "A2", "oops");
    engine.setCellFormula("Sheet1", "B5", "ROUND(A2,1)");
    expect(engine.getCellValue("Sheet1", "B5")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);
  });

  it("preserves topo order across mixed wasm and js formula runs", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 10);
    engine.setCellValue("Sheet1", "A2", 5);
    engine.setCellFormula("Sheet1", "B2", "A1+A2");
    engine.setCellFormula("Sheet1", "D1", "SUM(2:2)");

    engine.setCellValue("Sheet1", "A1", 12);

    expect(engine.getCellValue("Sheet1", "D1")).toEqual({ tag: ValueTag.Number, value: 22 });
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1);
    expect(engine.getLastMetrics().jsFormulaCount).toBe(1);
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

  it("rebinds formulas to #REF! when a referenced sheet is deleted", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.createSheet("Sheet2");
    engine.setCellValue("Sheet2", "B1", 3);
    engine.setCellFormula("Sheet1", "A1", "Sheet2!B1*2");

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 6 });

    engine.deleteSheet("Sheet2");

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });
  });

  it("clears reverse range edges when a range-backed formula is removed", async () => {
    const engine = new SpreadsheetEngine();
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellFormula("Sheet1", "B1", "SUM(A:A)");

    expect(engine.getDependencies("Sheet1", "A1").directDependents).toContain("Sheet1!B1");

    engine.clearCell("Sheet1", "B1");

    expect(engine.getDependencies("Sheet1", "A1").directDependents).toEqual([]);
  });

  it("rebinds column and row range formulas when new cells materialize later", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellFormula("Sheet1", "B1", "SUM(A:A)");
    engine.setCellFormula("Sheet1", "B3", "SUM(2:2)");

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "B3")).toEqual({ tag: ValueTag.Number, value: 0 });

    engine.setCellValue("Sheet1", "A4", 3);
    engine.setCellValue("Sheet1", "C2", 5);

    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(engine.getCellValue("Sheet1", "B3")).toEqual({ tag: ValueTag.Number, value: 5 });
    expect(engine.getDependencies("Sheet1", "B1").directPrecedents).toEqual(["Sheet1!A1", "Sheet1!A4"]);
    expect(engine.getDependencies("Sheet1", "B3").directPrecedents).toEqual(["Sheet1!C2"]);

    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    expect(b1Index).toBeDefined();
    const runtimeFormula = b1Index === undefined ? undefined : (engine as any).formulas.get(b1Index);
    expect(runtimeFormula?.dependencyIndices).toBeInstanceOf(Uint32Array);
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

  it("ignores duplicate remote batches and stale cell replays behind sheet tombstones", async () => {
    const primary = new SpreadsheetEngine({ workbookName: "spec", replicaId: "a" });
    const replica = new SpreadsheetEngine({ workbookName: "spec", replicaId: "b" });
    await Promise.all([primary.ready(), replica.ready()]);

    const outbound: EngineOpBatch[] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Sheet1");
    const createBatch = outbound[outbound.length - 1]!;

    primary.setCellValue("Sheet1", "A1", 7);
    const valueBatch = outbound[outbound.length - 1]!;

    replica.applyRemoteBatch(createBatch);
    replica.applyRemoteBatch(valueBatch);
    const versionBeforeDuplicate = replica.explainCell("Sheet1", "A1").version;

    replica.applyRemoteBatch(valueBatch);
    expect(replica.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(replica.explainCell("Sheet1", "A1").version).toBe(versionBeforeDuplicate);

    primary.deleteSheet("Sheet1");
    const deleteBatch = outbound[outbound.length - 1]!;
    replica.applyRemoteBatch(deleteBatch);
    expect(replica.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });

    const restored = new SpreadsheetEngine({ workbookName: "restored", replicaId: "b" });
    await restored.ready();
    restored.importSnapshot(replica.exportSnapshot());
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot());

    restored.applyRemoteBatch(valueBatch);
    expect(restored.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });
  });

  it("exports sparse high-row cells without truncating the sheet", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A10002", 7);

    const snapshot = engine.exportSnapshot();
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: "A10002", value: 7 });
  });

  it("roundtrips a single sheet through CSV with formulas and quoted strings", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 12);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setCellValue("Sheet1", "A2", "alpha,beta");

    const csv = engine.exportSheetCsv("Sheet1");
    expect(csv).toBe('12,=A1*2\n"alpha,beta",');

    const restored = new SpreadsheetEngine({ workbookName: "restored" });
    await restored.ready();
    restored.importSheetCsv("Sheet1", csv);

    expect(restored.getCell("Sheet1", "A1").value).toEqual({ tag: ValueTag.Number, value: 12 });
    expect(restored.getCell("Sheet1", "B1").formula).toBe("A1*2");
    expect(restored.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 24 });
    expect(restored.getCell("Sheet1", "A2").value).toEqual({ tag: ValueTag.String, value: "alpha,beta", stringId: 1 });
  });

  it("persists cell formats through imperative updates and snapshot roundtrip", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 12);
    engine.setCellFormat("Sheet1", "A1", "currency-usd");

    expect(engine.getCell("Sheet1", "A1").format).toBe("currency-usd");
    expect(engine.explainCell("Sheet1", "A1").format).toBe("currency-usd");

    const restored = new SpreadsheetEngine({ workbookName: "restored" });
    await restored.ready();
    restored.importSnapshot(engine.exportSnapshot());

    expect(restored.getCell("Sheet1", "A1").format).toBe("currency-usd");
    expect(restored.exportSnapshot().sheets[0]?.cells).toContainEqual({
      address: "A1",
      value: 12,
      format: "currency-usd"
    });
  });

  it("includes format-only mutations in changed cell events", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 12);

    const changed: number[][] = [];
    const unsubscribe = engine.subscribe((event) => {
      changed.push(Array.from(event.changedCellIndices));
    });

    engine.setCellFormat("Sheet1", "A1", "currency-usd");

    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    expect(a1Index).toBeDefined();
    expect(changed.at(-1)).toEqual([a1Index!]);

    unsubscribe();
  });

  it("replaces existing sheet contents on CSV import", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    engine.setCellValue("Sheet1", "C3", 9);

    engine.importSheetCsv("Sheet1", "7,8");

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 7 });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 8 });
    expect(engine.getCellValue("Sheet1", "C3")).toEqual({ tag: ValueTag.Empty });
  });

  it("explains formula cells with mode, version, and dependencies", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 5);
    engine.setCellFormula("Sheet1", "B1", "A1*2");

    const explanation = engine.explainCell("Sheet1", "B1");

    expect(explanation.formula).toBe("A1*2");
    expect(explanation.mode).toBeDefined();
    expect(explanation.version).toBeGreaterThan(0);
    expect(explanation.directPrecedents).toEqual(["Sheet1!A1"]);
    expect(explanation.directDependents).toEqual([]);
    expect(explanation.inCycle).toBe(false);
  });

  it("stores runtime formula metadata with real formula ids and dependency slices", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellFormula("Sheet1", "C3", "A1*2");

    const cellIndex = engine.workbook.getCellIndex("Sheet1", "C3");
    expect(cellIndex).toBeDefined();

    const formulaId = engine.workbook.cellStore.formulaIds[cellIndex!]!;
    const runtimeFormula = (
      engine as unknown as {
        formulas: {
          get(cellIndex: number): {
            compiled: {
              id: number;
              depsPtr: number;
              depsLen: number;
              programOffset: number;
              programLength: number;
            };
            dependencyEntities: { ptr: number; len: number };
            runtimeProgram: Uint32Array;
          } | undefined;
        };
      }
    ).formulas.get(cellIndex!);

    expect(formulaId).toBeGreaterThan(0);
    expect(runtimeFormula).toBeDefined();
    expect(runtimeFormula?.compiled.id).toBe(formulaId);
    expect(runtimeFormula?.compiled.depsPtr).toBe(runtimeFormula?.dependencyEntities.ptr);
    expect(runtimeFormula?.compiled.depsLen).toBe(runtimeFormula?.dependencyEntities.len);
    expect(runtimeFormula?.compiled.programOffset).toBe(0);
    expect(runtimeFormula?.compiled.programLength).toBe(runtimeFormula?.runtimeProgram.length);
  });

  it("patches runtime cell and range operands from packed symbolic binding buffers", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "symbolic-spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellValue("Sheet1", "B1", 3);
    engine.setCellValue("Sheet1", "C1", 5);
    engine.setCellFormula("Sheet1", "D1", "SUM(A1:B1)+C1");

    const cellIndex = engine.workbook.getCellIndex("Sheet1", "D1");
    const c1Index = engine.workbook.getCellIndex("Sheet1", "C1");
    expect(cellIndex).toBeDefined();
    expect(c1Index).toBeDefined();

    const runtimeFormula = (
      engine as unknown as {
        formulas: {
          get(cellIndex: number): {
            rangeDependencies: Uint32Array;
            runtimeProgram: Uint32Array;
          } | undefined;
        };
      }
    ).formulas.get(cellIndex!);

    expect(runtimeFormula).toBeDefined();
    const pushCell = runtimeFormula?.runtimeProgram.find((instruction) => (instruction >>> 24) === Opcode.PushCell);
    const pushRange = runtimeFormula?.runtimeProgram.find((instruction) => (instruction >>> 24) === Opcode.PushRange);

    expect(pushCell).toBeDefined();
    expect(pushRange).toBeDefined();
    expect(pushCell! & 0x00ff_ffff).toBe(c1Index);
    expect(pushRange! & 0x00ff_ffff).toBe(runtimeFormula?.rangeDependencies[0]);
  });

  it("assigns deterministic cycle group ids for cyclic formulas", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "cycle-spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellFormula("Sheet1", "A1", "B1+1");
    engine.setCellFormula("Sheet1", "B1", "A1+1");

    const a1Index = engine.workbook.getCellIndex("Sheet1", "A1");
    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");

    expect(a1Index).toBeDefined();
    expect(b1Index).toBeDefined();
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle });
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBeGreaterThanOrEqual(0);
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBe(engine.workbook.cellStore.cycleGroupIds[b1Index!]);
  });

  it("assigns topo ranks through range-node dependents deterministically", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "topo-spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 2);
    engine.setCellFormula("Sheet1", "B1", "A1*2");
    engine.setCellFormula("Sheet1", "D1", "SUM(A1:B1)");

    const b1Index = engine.workbook.getCellIndex("Sheet1", "B1");
    const d1Index = engine.workbook.getCellIndex("Sheet1", "D1");

    expect(b1Index).toBeDefined();
    expect(d1Index).toBeDefined();
    expect(engine.workbook.cellStore.topoRanks[b1Index!]!).toBeLessThan(engine.workbook.cellStore.topoRanks[d1Index!]!);
  });

  it("notifies per-cell listeners only for the cells that changed", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");

    let a1Notifications = 0;
    let b1Notifications = 0;
    const unsubscribeA1 = engine.subscribeCell("Sheet1", "A1", () => {
      a1Notifications += 1;
    });
    const unsubscribeB1 = engine.subscribeCell("Sheet1", "B1", () => {
      b1Notifications += 1;
    });

    engine.setCellValue("Sheet1", "A1", 1);
    expect(a1Notifications).toBe(1);
    expect(b1Notifications).toBe(0);

    engine.setCellValue("Sheet1", "B1", 2);
    expect(a1Notifications).toBe(1);
    expect(b1Notifications).toBe(1);

    engine.setCellValue("Sheet1", "C1", 3);
    expect(a1Notifications).toBe(1);
    expect(b1Notifications).toBe(1);

    unsubscribeA1();
    unsubscribeB1();
  });

  it("notifies grouped watched cells only when one of them changes", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");

    let notifications = 0;
    const unsubscribe = engine.subscribeCells("Sheet1", ["A1", "A2"], () => {
      notifications += 1;
    });

    engine.setCellValue("Sheet1", "B1", 5);
    expect(notifications).toBe(0);

    engine.setCellValue("Sheet1", "A2", 8);
    expect(notifications).toBe(1);

    unsubscribe();
  });

  it("notifies watched cells when sheet deletion clears them", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 7);

    let notifications = 0;
    const unsubscribe = engine.subscribeCell("Sheet1", "A1", () => {
      notifications += 1;
    });

    engine.deleteSheet("Sheet1");

    expect(notifications).toBe(1);
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });

    unsubscribe();
  });

  it("tracks selection state inside the engine and notifies subscribers", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();

    const seen: string[] = [];
    const unsubscribe = engine.subscribeSelection(() => {
      const snapshot = engine.getSelectionState();
      seen.push(`${snapshot.sheetName}!${snapshot.address ?? "null"}`);
    });

    engine.setSelection("Sheet2", "B3");
    engine.setSelection("Sheet2", "B3");
    engine.setSelection("Sheet1", "A1");

    expect(engine.getSelectionState()).toEqual({
      sheetName: "Sheet1",
      address: "A1",
      anchorAddress: "A1",
      range: { startAddress: "A1", endAddress: "A1" },
      editMode: "idle"
    });
    expect(seen).toEqual(["Sheet2!B3", "Sheet1!A1"]);

    unsubscribe();
  });

  it("supports range mutation helpers and undo/redo over the same local apply path", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");

    engine.setRangeValues(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      [
        [1, 2],
        [3, 4]
      ]
    );
    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "B2")).toEqual({ tag: ValueTag.Number, value: 4 });

    engine.setRangeFormulas(
      { sheetName: "Sheet1", startAddress: "C1", endAddress: "C2" },
      [
        ["SUM(A1:B1)"],
        ["SUM(A2:B2)"]
      ]
    );
    expect(engine.getCellValue("Sheet1", "C1")).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 7 });

    engine.clearRange({ sheetName: "Sheet1", startAddress: "A2", endAddress: "B2" });
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 0 });

    engine.undo();
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Number, value: 3 });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 7 });

    engine.redo();
    expect(engine.getCellValue("Sheet1", "A2")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "C2")).toEqual({ tag: ValueTag.Number, value: 0 });
  });

  it("copies and fills rectangular ranges", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setRangeValues(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      [
        [1, 2],
        [3, 4]
      ]
    );

    engine.copyRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      { sheetName: "Sheet1", startAddress: "D1", endAddress: "E2" }
    );
    expect(engine.getCellValue("Sheet1", "D1")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "E2")).toEqual({ tag: ValueTag.Number, value: 4 });

    engine.fillRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B1" },
      { sheetName: "Sheet1", startAddress: "A4", endAddress: "D5" }
    );
    expect(engine.getCellValue("Sheet1", "A4")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "B4")).toEqual({ tag: ValueTag.Number, value: 2 });
    expect(engine.getCellValue("Sheet1", "C4")).toEqual({ tag: ValueTag.Number, value: 1 });
    expect(engine.getCellValue("Sheet1", "D5")).toEqual({ tag: ValueTag.Number, value: 2 });
  });

  it("tracks sync client connection state and forwards local batches", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "spec" });
    await engine.ready();

    const forwarded: EngineOpBatch[] = [];
    let connected = false;
    let disconnected = false;

    await engine.connectSyncClient({
      async connect({ setState }) {
        connected = true;
        setState("behind");
        return {
          send(batch) {
            forwarded.push(batch);
          },
          async disconnect() {
            disconnected = true;
          }
        };
      }
    });

    expect(connected).toBe(true);
    expect(engine.getSyncState()).toBe("behind");

    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 9);

    expect(forwarded).toHaveLength(2);
    expect(forwarded[1]?.ops).toEqual([
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 9 }
    ]);

    await engine.disconnectSyncClient();
    expect(disconnected).toBe(true);
    expect(engine.getSyncState()).toBe("local-only");
  });
});
