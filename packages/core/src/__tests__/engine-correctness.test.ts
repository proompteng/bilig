import { isDeepStrictEqual } from "node:util";
import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { formatAddress } from "@bilig/formula";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStylePatch,
  LiteralInput,
  WorkbookSnapshot,
} from "@bilig/protocol";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import { SpreadsheetEngine } from "../engine.js";
import { runProperty } from "@bilig/test-fuzz";

type CoreCorrectnessAction =
  | { kind: "values"; range: CellRangeRef; values: LiteralInput[][] }
  | { kind: "formula"; address: string; formula: string }
  | { kind: "style"; range: CellRangeRef; patch: CellStylePatch }
  | { kind: "format"; range: CellRangeRef; format: CellNumberFormatInput }
  | { kind: "clear"; range: CellRangeRef }
  | { kind: "fill"; source: CellRangeRef; target: CellRangeRef }
  | { kind: "insertRows"; start: number; count: number }
  | { kind: "deleteRows"; start: number; count: number }
  | { kind: "insertColumns"; start: number; count: number }
  | { kind: "deleteColumns"; start: number; count: number };

const sheetName = "Sheet1";

function toRangeRef(
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
  };
}

function buildValueMatrix(
  height: number,
  width: number,
  values: readonly LiteralInput[],
): LiteralInput[][] {
  const rows: LiteralInput[][] = [];
  let offset = 0;
  for (let row = 0; row < height; row += 1) {
    const nextRow: LiteralInput[] = [];
    for (let col = 0; col < width; col += 1) {
      nextRow.push(values[offset] ?? null);
      offset += 1;
    }
    rows.push(nextRow);
  }
  return rows;
}

function assertSnapshotInvariants(snapshot: WorkbookSnapshot): void {
  const sheetNames = snapshot.sheets.map((sheet) => sheet.name);
  expect(new Set(sheetNames).size).toBe(sheetNames.length);
  snapshot.sheets.forEach((sheet) => {
    const addresses = sheet.cells.map((cell) => cell.address);
    expect(new Set(addresses).size).toBe(addresses.length);
  });
}

async function createBaselineSnapshot(workbookName: string): Promise<WorkbookSnapshot> {
  const seed = new SpreadsheetEngine({
    workbookName,
    replicaId: `${workbookName}-seed`,
  });
  await seed.ready();
  seed.createSheet(sheetName);
  return seed.exportSnapshot();
}

function applyAction(engine: SpreadsheetEngine, action: CoreCorrectnessAction): void {
  switch (action.kind) {
    case "values":
      engine.setRangeValues(action.range, action.values);
      break;
    case "formula":
      engine.setCellFormula(sheetName, action.address, action.formula);
      break;
    case "style":
      engine.setRangeStyle(action.range, action.patch);
      break;
    case "format":
      engine.setRangeNumberFormat(action.range, action.format);
      break;
    case "clear":
      engine.clearRange(action.range);
      break;
    case "fill":
      engine.fillRange(action.source, action.target);
      break;
    case "insertRows":
      engine.insertRows(sheetName, action.start, action.count);
      break;
    case "deleteRows":
      engine.deleteRows(sheetName, action.start, action.count);
      break;
    case "insertColumns":
      engine.insertColumns(sheetName, action.start, action.count);
      break;
    case "deleteColumns":
      engine.deleteColumns(sheetName, action.start, action.count);
      break;
  }
}

function undoAll(engine: SpreadsheetEngine, maxSteps: number): number {
  let steps = 0;
  while (engine.undo()) {
    steps += 1;
    if (steps > maxSteps) {
      throw new Error(`Undo exceeded expected history budget: ${steps} > ${maxSteps}`);
    }
  }
  return steps;
}

function redoAll(engine: SpreadsheetEngine, maxSteps: number): number {
  let steps = 0;
  while (engine.redo()) {
    steps += 1;
    if (steps > maxSteps) {
      throw new Error(`Redo exceeded expected history budget: ${steps} > ${maxSteps}`);
    }
  }
  return steps;
}

const literalInputArbitrary = fc.oneof<LiteralInput>(
  fc.integer({ min: -10_000, max: 10_000 }),
  fc.boolean(),
  fc.constantFrom("north", "south", "ready", "done"),
  fc.constant(null),
);
const rangeSeedArbitrary = fc.record({
  startRow: fc.integer({ min: 0, max: 5 }),
  startCol: fc.integer({ min: 0, max: 5 }),
  height: fc.integer({ min: 1, max: 2 }),
  width: fc.integer({ min: 1, max: 2 }),
});
const rangeArbitrary = rangeSeedArbitrary.map((value) =>
  toRangeRef(
    value.startRow,
    value.startCol,
    value.startRow + value.height - 1,
    value.startCol + value.width - 1,
  ),
);
const formulaArbitrary = fc
  .tuple(
    fc.constantFrom("A1", "B2", "C3", "D4", "E5"),
    fc.constantFrom("+", "-", "*", "/"),
    fc.constantFrom("A1", "B2", "C3", "D4", "E5"),
  )
  .map(([left, operator, right]) => `${left}${operator}${right}`);
const stylePatchArbitrary = fc.constantFrom<CellStylePatch>(
  { fill: { backgroundColor: "#dbeafe" } },
  { font: { bold: true } },
  { alignment: { horizontal: "right", wrap: true } },
);
const formatInputArbitrary = fc.constantFrom<CellNumberFormatInput>(
  "0.00",
  { kind: "currency", currency: "USD", decimals: 2 },
  { kind: "percent", decimals: 1 },
  { kind: "text" },
);
const valuesActionArbitrary = rangeSeedArbitrary.chain((range) =>
  fc
    .array(literalInputArbitrary, {
      minLength: range.height * range.width,
      maxLength: range.height * range.width,
    })
    .map((values) => ({
      kind: "values" as const,
      range: toRangeRef(
        range.startRow,
        range.startCol,
        range.startRow + range.height - 1,
        range.startCol + range.width - 1,
      ),
      values: buildValueMatrix(range.height, range.width, values),
    })),
);
const formulaActionArbitrary = fc
  .record({
    row: fc.integer({ min: 0, max: 5 }),
    col: fc.integer({ min: 0, max: 5 }),
    formula: formulaArbitrary,
  })
  .map(({ row, col, formula }) => ({
    kind: "formula" as const,
    address: formatAddress(row, col),
    formula,
  }));
const styleActionArbitrary = fc
  .record({ range: rangeArbitrary, patch: stylePatchArbitrary })
  .map(({ range, patch }) => ({ kind: "style" as const, range, patch }));
const formatActionArbitrary = fc
  .record({ range: rangeArbitrary, format: formatInputArbitrary })
  .map(({ range, format }) => ({ kind: "format" as const, range, format }));
const clearActionArbitrary = rangeArbitrary.map((range) => ({ kind: "clear" as const, range }));
const fillActionArbitrary = rangeSeedArbitrary.chain((source) =>
  fc
    .record({
      targetStartRow: fc.integer({ min: source.startRow, max: 5 }),
      targetStartCol: fc.integer({ min: source.startCol, max: 5 }),
    })
    .map(({ targetStartRow, targetStartCol }) => ({
      kind: "fill" as const,
      source: toRangeRef(
        source.startRow,
        source.startCol,
        source.startRow + source.height - 1,
        source.startCol + source.width - 1,
      ),
      target: toRangeRef(
        targetStartRow,
        targetStartCol,
        Math.min(5, targetStartRow + source.height - 1),
        Math.min(5, targetStartCol + source.width - 1),
      ),
    })),
);
const axisMutationArbitrary = fc.record({
  start: fc.integer({ min: 0, max: 4 }),
  count: fc.integer({ min: 1, max: 2 }),
});
const insertRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "insertRows" as const,
  start,
  count,
}));
const deleteRowsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "deleteRows" as const,
  start,
  count,
}));
const insertColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "insertColumns" as const,
  start,
  count,
}));
const deleteColumnsActionArbitrary = axisMutationArbitrary.map(({ start, count }) => ({
  kind: "deleteColumns" as const,
  start,
  count,
}));
const correctnessActionArbitrary = fc.oneof<CoreCorrectnessAction>(
  valuesActionArbitrary,
  formulaActionArbitrary,
  styleActionArbitrary,
  formatActionArbitrary,
  clearActionArbitrary,
  fillActionArbitrary,
  insertRowsActionArbitrary,
  deleteRowsActionArbitrary,
  insertColumnsActionArbitrary,
  deleteColumnsActionArbitrary,
);

describe("engine correctness", () => {
  it("clears sparse style and format metadata when undoing structural edit sequences", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-undo-sparse-ranges");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-undo-sparse-ranges",
      replicaId: "correctness-undo-sparse-ranges",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.insertColumns(sheetName, 0, 1);
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );
    engine.deleteColumns(sheetName, 0, 1);
    engine.setRangeNumberFormat({ sheetName, startAddress: "A1", endAddress: "A1" }, "0.00");

    expect(undoAll(engine, 16)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("prunes materialized empty cells when undo restores a blank address", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-undo-empty-cell-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-undo-empty-cell-prune",
      replicaId: "correctness-undo-empty-cell-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "A1" }, [[false]]);
    engine.insertRows(sheetName, 0, 1);
    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "A1" }, [[0]]);
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );

    expect(undoAll(engine, 16)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("replays inserted column identities exactly across undo and redo", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-redo-column-identity");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-redo-column-identity",
      replicaId: "correctness-redo-column-identity",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setCellFormula(sheetName, "A1", "A1+A1");
    engine.insertColumns(sheetName, 0, 1);
    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "A1" }, [[null]]);
    engine.setRangeNumberFormat({ sheetName, startAddress: "A1", endAddress: "A1" }, "0.00");

    const finalSnapshot = engine.exportSnapshot();
    expect(finalSnapshot.sheets[0]?.metadata?.columns).toEqual([{ id: "column-1", index: 0 }]);

    const undoCount = undoAll(engine, 16);
    expect(undoCount).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);

    const redoCount = redoAll(engine, undoCount + 2);
    expect(redoCount).toBe(undoCount);
    expect(engine.exportSnapshot()).toEqual(finalSnapshot);
  });

  it("does not leave empty cells behind when fill replays a blank source cell", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-fill-empty-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-fill-empty-prune",
      replicaId: "correctness-fill-empty-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setCellFormula(sheetName, "B1", "A1+A1");
    engine.deleteColumns(sheetName, 0, 1);
    engine.fillRange(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { sheetName, startAddress: "A1", endAddress: "A1" },
    );
    engine.clearRange({ sheetName, startAddress: "A1", endAddress: "A1" });

    expect(undoAll(engine, 16)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("prunes translated dependency cells after structural formula rebuild undo", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-structural-dependency-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-dependency-prune",
      replicaId: "correctness-structural-dependency-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setCellFormula(sheetName, "B1", "A1+D4");
    engine.deleteRows(sheetName, 1, 1);
    engine.insertRows(sheetName, 0, 1);
    engine.insertRows(sheetName, 0, 1);
    engine.deleteRows(sheetName, 3, 2);

    expect(undoAll(engine, 20)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("does not preserve temporary null dependency placeholders as authored blanks during undo", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-undo-temporary-blank-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-undo-temporary-blank-prune",
      replicaId: "correctness-undo-temporary-blank-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setCellFormula(sheetName, "A1", "C3+A1");
    engine.setRangeValues({ sheetName, startAddress: "B3", endAddress: "C3" }, [[0, false]]);
    engine.deleteRows(sheetName, 0, 1);
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );
    engine.setCellFormula(sheetName, "A1", "A1+A1");
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );

    expect(undoAll(engine, 24)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("does not reify inherited number formats into explicit cells during structural undo", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-structural-format-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-format-prune",
      replicaId: "correctness-structural-format-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setRangeValues({ sheetName, startAddress: "A3", endAddress: "A3" }, [[0]]);
    engine.setRangeNumberFormat({ sheetName, startAddress: "A3", endAddress: "A3" }, "0.00");
    engine.clearRange({ sheetName, startAddress: "A1", endAddress: "A1" });
    engine.deleteRows(sheetName, 0, 1);
    engine.deleteColumns(sheetName, 0, 1);
    engine.clearRange({ sheetName, startAddress: "A1", endAddress: "A1" });

    expect(undoAll(engine, 24)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("does not materialize inherited format-range placeholders during snapshot roundtrip", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-snapshot-format-range",
      replicaId: "correctness-snapshot-format-range",
    });
    await engine.ready();
    engine.createSheet(sheetName);

    engine.setCellFormula(sheetName, "A1", "C4+C3");
    engine.setRangeNumberFormat({ sheetName, startAddress: "B4", endAddress: "C4" }, "0.00");
    engine.clearRange({ sheetName, startAddress: "B4", endAddress: "C4" });

    const snapshot = engine.exportSnapshot();
    const restored = new SpreadsheetEngine({
      workbookName: snapshot.workbook.name,
      replicaId: "correctness-snapshot-format-range-restored",
    });
    await restored.ready();
    restored.importSnapshot(snapshot);

    expect(restored.exportSnapshot()).toEqual(snapshot);
  });

  it("preserves rewritten error formulas across structural delete undo", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-formula-error-undo",
      replicaId: "correctness-structural-formula-error-undo",
    });
    await engine.ready();
    engine.createSheet(sheetName);

    engine.setCellFormula(sheetName, "A1", "A1+D4");
    engine.deleteRows(sheetName, 2, 2);

    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "correctness-structural-formula-error-undo" },
      sheets: [
        {
          id: 1,
          name: sheetName,
          order: 0,
          cells: [{ address: "A1", formula: "A1+#REF!" }],
        },
      ],
    });

    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toEqual({
      version: 1,
      workbook: { name: "correctness-structural-formula-error-undo" },
      sheets: [
        {
          id: 1,
          name: sheetName,
          order: 0,
          cells: [{ address: "A1", formula: "A1+D4" }],
        },
      ],
    });
  });

  it("restores metadata-only number formats after structural delete undo", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-format-only-undo",
      replicaId: "correctness-structural-format-only-undo",
    });
    await engine.ready();
    engine.createSheet(sheetName);

    engine.setRangeNumberFormat({ sheetName, startAddress: "A1", endAddress: "A1" }, "0.00");
    const initialSnapshot = engine.exportSnapshot();

    engine.deleteColumns(sheetName, 0, 1);
    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("restores named-range structures after insert-column undo", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-named-range-undo",
      replicaId: "correctness-structural-named-range-undo",
    });
    await engine.ready();
    engine.createSheet(sheetName);
    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "B3" }, [
      ["Qty", "Amount"],
      [1, 10],
      [2, 20],
    ]);
    engine.setDefinedName("SalesRange", {
      kind: "range-ref",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
    });
    engine.setTable({
      name: "Sales",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });
    engine.setCellFormula(sheetName, "C1", "SUM(SalesRange)");

    const initialSnapshot = engine.exportSnapshot();

    engine.insertColumns(sheetName, 0, 1);
    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName,
        startAddress: "B1",
        endAddress: "C3",
      },
    });
    expect(engine.getTable("Sales")).toEqual({
      name: "Sales",
      sheetName,
      startAddress: "B1",
      endAddress: "C3",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });
    expect(engine.getCell("Sheet1", "D1").formula).toBe("SUM(SalesRange)");
    expect(engine.getCellValue(sheetName, "D1")).toMatchObject({ tag: 1, value: 33 });

    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName,
        startAddress: "A1",
        endAddress: "B3",
      },
    });
    expect(engine.getTable("Sales")).toEqual({
      name: "Sales",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });
    expect(engine.getCell(sheetName, "C1").formula).toBe("SUM(SalesRange)");
    expect(engine.getCellValue(sheetName, "C1")).toMatchObject({ tag: 1, value: 33 });
  });

  it("restores named-range structures after delete-row undo", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-named-range-delete-row-undo",
      replicaId: "correctness-structural-named-range-delete-row-undo",
    });
    await engine.ready();
    engine.createSheet(sheetName);
    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "B3" }, [
      ["Qty", "Amount"],
      [1, 10],
      [2, 20],
    ]);
    engine.setDefinedName("SalesRange", {
      kind: "range-ref",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
    });
    engine.setTable({
      name: "Sales",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });
    engine.setCellFormula(sheetName, "C1", "SUM(SalesRange)");

    const initialSnapshot = engine.exportSnapshot();

    engine.deleteRows(sheetName, 2, 2);
    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName,
        startAddress: "A1",
        endAddress: "B2",
      },
    });
    expect(engine.getTable("Sales")).toEqual({
      name: "Sales",
      sheetName,
      startAddress: "A1",
      endAddress: "B2",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });

    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
    expect(engine.getDefinedName("SalesRange")).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName,
        startAddress: "A1",
        endAddress: "B3",
      },
    });
    expect(engine.getTable("Sales")).toEqual({
      name: "Sales",
      sheetName,
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Qty", "Amount"],
      headerRow: true,
      totalsRow: false,
    });
    expect(engine.getCell(sheetName, "C1").formula).toBe("SUM(SalesRange)");
    expect(engine.getCellValue(sheetName, "C1")).toMatchObject({ tag: 1, value: 33 });
  });

  it("restores filter, sort, and validation metadata after delete-row undo", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-structural-sheet-metadata-delete-row-undo",
      replicaId: "correctness-structural-sheet-metadata-delete-row-undo",
    });
    await engine.ready();
    engine.createSheet(sheetName);
    engine.setRangeValues({ sheetName, startAddress: "A1", endAddress: "C3" }, [
      ["Id", "Status", "Amount"],
      [1, "Draft", 10],
      [2, "Final", 20],
    ]);
    engine.setFilter(sheetName, { sheetName, startAddress: "A1", endAddress: "C3" });
    engine.setSort(sheetName, { sheetName, startAddress: "A1", endAddress: "C3" }, [
      { keyAddress: "B1", direction: "asc" },
    ]);
    engine.setDataValidation({
      range: { sheetName, startAddress: "B2", endAddress: "B3" },
      rule: { kind: "list", values: ["Draft", "Final"] },
      allowBlank: false,
      showDropdown: true,
    });

    const initialSnapshot = engine.exportSnapshot();

    engine.deleteRows(sheetName, 0, 1);
    expect(engine.getFilters(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: "A1", endAddress: "C2" },
      },
    ]);

    expect(engine.undo()).toBe(true);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
    expect(engine.getFilters(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: "A1", endAddress: "C3" },
      },
    ]);
    expect(engine.getSorts(sheetName)).toEqual([
      {
        sheetName,
        range: { sheetName, startAddress: "A1", endAddress: "C3" },
        keys: [{ keyAddress: "B1", direction: "asc" }],
      },
    ]);
    expect(engine.getDataValidations(sheetName)).toEqual([
      {
        range: { sheetName, startAddress: "B2", endAddress: "B3" },
        rule: { kind: "list", values: ["Draft", "Final"] },
        allowBlank: false,
        showDropdown: true,
      },
    ]);
  });

  it("prunes orphaned explicit formats after structural undo restores", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-undo-format-orphan-prune");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-undo-format-orphan-prune",
      replicaId: "correctness-undo-format-orphan-prune",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setRangeNumberFormat({ sheetName, startAddress: "A1", endAddress: "B2" }, "0.00");
    engine.insertColumns(sheetName, 0, 1);
    engine.insertColumns(sheetName, 0, 1);
    engine.setCellFormula(sheetName, "A1", "D4+A1");
    engine.deleteColumns(sheetName, 1, 1);
    engine.fillRange(
      { sheetName, startAddress: "B1", endAddress: "C2" },
      { sheetName, startAddress: "B4", endAddress: "C5" },
    );
    engine.setRangeNumberFormat({ sheetName, startAddress: "A1", endAddress: "A1" }, "0.00");
    engine.insertColumns(sheetName, 0, 1);
    engine.deleteRows(sheetName, 3, 1);
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );

    expect(undoAll(engine, 24)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
  });

  it("drops orphaned formula dependents during structural undo replay", async () => {
    const initialSnapshot = await createBaselineSnapshot("correctness-orphan-formula-undo");
    const engine = new SpreadsheetEngine({
      workbookName: "correctness-orphan-formula-undo",
      replicaId: "correctness-orphan-formula-undo",
    });
    await engine.ready();
    engine.importSnapshot(initialSnapshot);

    engine.setCellFormula(sheetName, "A1", "A1+B2");
    engine.setCellFormula(sheetName, "B3", "A1+A1");
    engine.setRangeStyle(
      { sheetName, startAddress: "A1", endAddress: "A1" },
      { fill: { backgroundColor: "#dbeafe" } },
    );
    engine.deleteRows(sheetName, 1, 1);
    engine.insertRows(sheetName, 0, 1);
    engine.clearRange({ sheetName, startAddress: "A1", endAddress: "A1" });

    expect(undoAll(engine, 24)).toBeGreaterThan(0);
    expect(engine.exportSnapshot()).toEqual(initialSnapshot);
    expect(engine.getDependents(sheetName, "A1")).toEqual({
      directPrecedents: [],
      directDependents: [],
    });
  });

  it("reverses random local edit streams through undo and redo", async () => {
    await runProperty({
      suite: "core/undo-redo-reversibility",
      arbitrary: fc.array(correctnessActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const initialSnapshot = await createBaselineSnapshot("correctness-undo-redo");
        const engine = new SpreadsheetEngine({
          workbookName: "correctness-undo-redo",
          replicaId: "correctness-undo-redo",
        });
        await engine.ready();
        engine.importSnapshot(initialSnapshot);

        let observedSemanticChange = false;
        for (const action of actions) {
          applyAction(engine, action);
          const currentSnapshot = engine.exportSnapshot();
          assertSnapshotInvariants(currentSnapshot);
          observedSemanticChange ||= !isDeepStrictEqual(currentSnapshot, initialSnapshot);
        }

        const finalSnapshot = engine.exportSnapshot();
        const undoCount = undoAll(engine, actions.length * 4);
        expect(undoCount > 0).toBe(observedSemanticChange);
        expect(engine.exportSnapshot()).toEqual(initialSnapshot);

        const redoCount = redoAll(engine, undoCount + 2);
        expect(redoCount).toBe(undoCount);
        expect(engine.exportSnapshot()).toEqual(finalSnapshot);
      },
    });
  });

  it("replays captured local batches into an equivalent replica state", async () => {
    await runProperty({
      suite: "core/local-batch-replay-parity",
      arbitrary: fc.array(correctnessActionArbitrary, { minLength: 4, maxLength: 18 }),
      predicate: async (actions) => {
        const initialSnapshot = await createBaselineSnapshot("correctness-replay");
        const primary = new SpreadsheetEngine({
          workbookName: "correctness-replay",
          replicaId: "primary",
        });
        const replica = new SpreadsheetEngine({
          workbookName: "correctness-replay",
          replicaId: "replica",
        });
        await Promise.all([primary.ready(), replica.ready()]);

        const outbound: EngineOpBatch[] = [];
        primary.subscribeBatches((batch) => outbound.push(batch));

        primary.importSnapshot(initialSnapshot);
        replica.importSnapshot(initialSnapshot);

        let appliedBatches = 0;
        expect(replica.exportSnapshot()).toEqual(primary.exportSnapshot());

        for (const action of actions) {
          applyAction(primary, action);
          while (appliedBatches < outbound.length) {
            const nextBatch = outbound[appliedBatches];
            if (!nextBatch) {
              throw new Error(`Missing outbound batch at index ${appliedBatches}`);
            }
            replica.applyRemoteBatch(nextBatch);
            appliedBatches += 1;
          }
          const primarySnapshot = primary.exportSnapshot();
          assertSnapshotInvariants(primarySnapshot);
          expect(replica.exportSnapshot()).toEqual(primarySnapshot);
        }
      },
    });
  });
});
