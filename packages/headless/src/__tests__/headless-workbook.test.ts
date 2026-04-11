import { afterEach, describe, expect, it } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";

import {
  HeadlessEvaluationSuspendedError,
  HeadlessWorkbook,
  type HeadlessCellAddress,
} from "../index.js";

const TEST_LANGUAGE_CODE = "xHF";

function cell(sheet: number, row: number, col: number): HeadlessCellAddress {
  return { sheet, row, col };
}

afterEach(() => {
  HeadlessWorkbook.unregisterAllFunctions();
  if (HeadlessWorkbook.getRegisteredLanguagesCodes().includes(TEST_LANGUAGE_CODE)) {
    HeadlessWorkbook.unregisterLanguage(TEST_LANGUAGE_CODE);
  }
});

describe("HeadlessWorkbook", () => {
  it("builds from named sheets and exposes stable sheet ids and serialization helpers", () => {
    const workbook = HeadlessWorkbook.buildFromSheets({
      Summary: [[1, "=A1*2"]],
      Detail: [[3]],
    });

    const summaryId = workbook.getSheetId("Summary")!;

    expect(workbook.getSheetName(summaryId)).toBe("Summary");
    expect(workbook.countSheets()).toBe(2);
    expect(workbook.getCellValue(cell(summaryId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(workbook.getCellFormula(cell(summaryId, 0, 1))).toBe("=A1*2");
    expect(workbook.getCellSerialized(cell(summaryId, 0, 1))).toBe("=A1*2");
    expect(workbook.getSheetDimensions(summaryId)).toEqual({ width: 2, height: 1 });
    expect(workbook.simpleCellAddressFromString("Summary!B1")).toEqual(cell(summaryId, 0, 1));
    expect(workbook.simpleCellRangeFromString("Summary!A1:B1")).toEqual({
      start: cell(summaryId, 0, 0),
      end: cell(summaryId, 0, 1),
    });
  });

  it("keeps literal-only initialization compatible with named expressions and later formulas", () => {
    const workbook = HeadlessWorkbook.buildFromSheets(
      {
        Bench: [
          [2, "west", true],
          [4, null, false],
        ],
      },
      {},
      [{ name: "BenchTotal", expression: "=SUM(Bench!$A$1:$A$2)" }],
    );
    const sheetId = workbook.getSheetId("Bench")!;

    expect(workbook.getNamedExpressionValue("BenchTotal")).toEqual({
      tag: ValueTag.Number,
      value: 6,
    });

    const changes = workbook.setCellContents(cell(sheetId, 0, 3), "=BenchTotal+A1");

    expect(changes).toHaveLength(1);
    expect(workbook.getCellValue(cell(sheetId, 0, 3))).toEqual({
      tag: ValueTag.Number,
      value: 8,
    });
    expect(workbook.getSheetSerialized(sheetId)).toEqual([
      [2, "west", true, "=BenchTotal+A1"],
      [4, null, false, null],
    ]);
  });

  it("supports sheet-scoped named expressions and restores public formulas", () => {
    const workbook = HeadlessWorkbook.buildFromSheets({
      Summary: [[]],
      Detail: [[]],
    });
    const summaryId = workbook.getSheetId("Summary")!;
    const detailId = workbook.getSheetId("Detail")!;
    const events: string[] = [];

    workbook.on("namedExpressionAdded", (name, changes) => {
      events.push(`add:${name}:${changes.length}`);
    });
    workbook.onDetailed("namedExpressionAdded", (payload) => {
      events.push(`scope:${payload.scope}`);
    });
    workbook.on("valuesUpdated", (changes) => {
      events.push(`values:${changes.length}`);
    });

    workbook.addNamedExpression("Rate", "=1", summaryId);
    workbook.addNamedExpression("Rate", "=2", detailId);

    expect(workbook.setCellContents(cell(summaryId, 0, 0), "=Rate+1")).toHaveLength(1);
    expect(workbook.setCellContents(cell(detailId, 0, 0), "=Rate+1")).toHaveLength(1);

    expect(workbook.getCellValue(cell(summaryId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(workbook.getCellValue(cell(detailId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 3,
    });
    expect(workbook.getCellFormula(cell(summaryId, 0, 0))).toBe("=Rate+1");
    expect(workbook.getCellFormula(cell(detailId, 0, 0))).toBe("=Rate+1");
    expect(workbook.getNamedExpressionValue("Rate", summaryId)).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(workbook.getNamedExpressionValue("Rate", detailId)).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(events.slice(0, 6)).toEqual([
      "add:Rate:1",
      "scope:1",
      "values:1",
      "add:Rate:1",
      "scope:2",
      "values:1",
    ]);
  });

  it("coalesces batch history into one undo entry and emits one values update", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[1]]);
    const sheetId = workbook.getSheetId("Sheet1")!;
    const valuesUpdated: number[] = [];
    const nestedMutationResults: number[] = [];

    workbook.on("valuesUpdated", (changes) => {
      valuesUpdated.push(changes.length);
    });

    const changes = workbook.batch(() => {
      nestedMutationResults.push(workbook.setCellContents(cell(sheetId, 0, 1), "=A1*2").length);
      nestedMutationResults.push(workbook.setCellContents(cell(sheetId, 1, 0), 5).length);
    });

    expect(changes).toHaveLength(2);
    expect(nestedMutationResults).toEqual([0, 0]);
    expect(valuesUpdated).toEqual([2]);
    expect(workbook.isThereSomethingToUndo()).toBe(true);

    const undoChanges = workbook.undo();

    expect(undoChanges).toHaveLength(2);
    expect(workbook.getCellValue(cell(sheetId, 0, 1)).tag).toBe(ValueTag.Empty);
    expect(workbook.getCellValue(cell(sheetId, 1, 0)).tag).toBe(ValueTag.Empty);
  });

  it("flushes deferred literal edits before formula writes inside a batch", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[1]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    const changes = workbook.batch(() => {
      expect(workbook.setCellContents(cell(sheetId, 0, 0), 10)).toEqual([]);
      expect(workbook.setCellContents(cell(sheetId, 0, 1), "=A1*2")).toEqual([]);
    });

    expect(changes).toHaveLength(2);
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    });
  });

  it("undoes and redoes deferred literal-only batches", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[1], [2]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 0), 10);
      workbook.setCellContents(cell(sheetId, 1, 0), 20);
    });

    expect(changes).toHaveLength(2);
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    });
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    });

    workbook.undo();
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });

    workbook.redo();
    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    });
    expect(workbook.getCellValue(cell(sheetId, 1, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 20,
    });
  });

  it("keeps exact MATCH correct when useColumnIndex is enabled", () => {
    const workbook = HeadlessWorkbook.buildFromSheets(
      {
        Bench: [[1, "", "", 2, "=MATCH(D1,A1:A3,0)"], [2], [3]],
      },
      { useColumnIndex: true },
    );
    const sheetId = workbook.getSheetId("Bench")!;

    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });

    const missingMatchChanges = workbook.setCellContents(cell(sheetId, 1, 0), 20);
    expect(
      missingMatchChanges.map((change) => (change.kind === "cell" ? change.a1 : change.kind)),
    ).toEqual(["E1", "A2"]);
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });

    const restoredMatchChanges = workbook.setCellContents(cell(sheetId, 0, 3), 3);
    expect(
      restoredMatchChanges.map((change) => (change.kind === "cell" ? change.a1 : change.kind)),
    ).toEqual(["D1", "E1"]);
    expect(workbook.getCellValue(cell(sheetId, 0, 4))).toMatchObject({
      tag: ValueTag.Number,
      value: 3,
    });
  });

  it("replaces literal sheet content in one undoable batch, including clears", () => {
    const workbook = HeadlessWorkbook.buildFromArray([
      [1, 2],
      [3, 4],
    ]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    const changes = workbook.setSheetContent(sheetId, [
      [10, 20],
      [null, 5],
    ]);

    expect(changes).toHaveLength(4);
    expect(workbook.getCellSerialized(cell(sheetId, 0, 0))).toBe(10);
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe(20);
    expect(workbook.getCellSerialized(cell(sheetId, 1, 0))).toBeNull();
    expect(workbook.getCellSerialized(cell(sheetId, 1, 1))).toBe(5);

    const undoChanges = workbook.undo();

    expect(undoChanges).toHaveLength(4);
    expect(workbook.getSheetSerialized(sheetId)).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it("suppresses readable value getters while evaluation is suspended and flushes on resume", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[1]]);
    const sheetId = workbook.getSheetId("Sheet1")!;
    const events: string[] = [];

    workbook.on("evaluationSuspended", () => {
      events.push("suspend");
    });
    workbook.on("evaluationResumed", (changes) => {
      events.push(`resume:${changes.length}`);
    });
    workbook.on("valuesUpdated", (changes) => {
      events.push(`values:${changes.length}`);
    });

    workbook.suspendEvaluation();
    workbook.setCellContents(cell(sheetId, 0, 1), "=A1+1");

    expect(() => workbook.getCellValue(cell(sheetId, 0, 1))).toThrow(
      HeadlessEvaluationSuspendedError,
    );

    const changes = workbook.resumeEvaluation();

    expect(changes).toHaveLength(1);
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(events).toEqual(["suspend", "resume:1", "values:1"]);
  });

  it("supports custom scalar functions and clipboard translation for pasted formulas", () => {
    HeadlessWorkbook.registerFunctionPlugin({
      id: "custom-math",
      implementedFunctions: {
        DOUBLE: { method: "DOUBLE" },
      },
      functions: {
        DOUBLE: (value) => {
          if (value?.tag !== ValueTag.Number) {
            return { tag: ValueTag.Error, code: 3 };
          }
          return { tag: ValueTag.Number, value: value.value * 2 };
        },
      },
    });

    const workbook = HeadlessWorkbook.buildFromArray([[2]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    workbook.setCellContents(cell(sheetId, 0, 1), "=DOUBLE(A1)");

    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    });
    expect(workbook.calculateFormula("=DOUBLE(3)")).toMatchObject({
      tag: ValueTag.Number,
      value: 6,
    });

    const copied = workbook.copy({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 0, 1),
    });
    expect(copied[0]?.[1]).toMatchObject({ tag: ValueTag.Number, value: 4 });

    workbook.paste(cell(sheetId, 1, 0));

    expect(workbook.getCellSerialized(cell(sheetId, 1, 0))).toBe(2);
    expect(workbook.getCellFormula(cell(sheetId, 1, 1))).toBe("=DOUBLE(A2)");
    expect(workbook.getCellValue(cell(sheetId, 1, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    });
  });

  it("rebuilds engine state when config changes affect available function plugins", () => {
    const plugin = {
      id: "custom-math",
      implementedFunctions: {
        DOUBLE: { method: "DOUBLE" },
      },
      functions: {
        DOUBLE: (value) => {
          if (value?.tag !== ValueTag.Number) {
            return { tag: ValueTag.Error, code: ErrorCode.Value };
          }
          return { tag: ValueTag.Number, value: value.value * 2 };
        },
      },
    } as const;

    HeadlessWorkbook.registerFunctionPlugin(plugin);

    const workbook = HeadlessWorkbook.buildFromArray([[2]], { functionPlugins: [plugin] });
    const sheetId = workbook.getSheetId("Sheet1")!;

    workbook.setCellContents(cell(sheetId, 0, 1), "=DOUBLE(A1)");
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 4,
    });

    workbook.updateConfig({
      functionPlugins: [{ id: "missing-plugin", implementedFunctions: {} }],
    });

    expect(workbook.getCellFormula(cell(sheetId, 0, 1))).toBe("=DOUBLE(A1)");
    expect(workbook.getCellValue(cell(sheetId, 0, 1))).toMatchObject({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    });
  });

  it("preserves workbook semantics across rebuildAndRecalculate and non-semantic config rebuilds", () => {
    const workbook = HeadlessWorkbook.buildFromSheets(
      {
        Data: [[1, "=A1*Rate"], [3], [5]],
        Summary: [["=FILTER(Data!A1:A3,Data!A1:A3>2)"]],
      },
      {
        useStats: false,
      },
    );
    const dataId = workbook.getSheetId("Data")!;
    const summaryId = workbook.getSheetId("Summary")!;

    workbook.addNamedExpression("Rate", "=2");

    const beforeDataSerialized = workbook.getSheetSerialized(dataId);
    const beforeSummaryValues = workbook.getRangeValues({
      start: cell(summaryId, 0, 0),
      end: cell(summaryId, 1, 0),
    });
    const beforeRateValue = workbook.getNamedExpressionValue("Rate");
    const rebuildChanges = workbook.rebuildAndRecalculate();

    expect(rebuildChanges).toEqual([]);
    expect(workbook.getSheetSerialized(dataId)).toEqual(beforeDataSerialized);
    expect(
      workbook.getRangeValues({
        start: cell(summaryId, 0, 0),
        end: cell(summaryId, 1, 0),
      }),
    ).toEqual(beforeSummaryValues);
    expect(workbook.getNamedExpressionValue("Rate")).toEqual(beforeRateValue);

    workbook.updateConfig({
      useColumnIndex: true,
      useStats: true,
    });

    expect(workbook.getCellFormula(cell(dataId, 0, 1))).toBe("=A1*Rate");
    expect(workbook.getCellValue(cell(dataId, 0, 1))).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(
      workbook.getRangeValues({
        start: cell(summaryId, 0, 0),
        end: cell(summaryId, 1, 0),
      }),
    ).toEqual(beforeSummaryValues);
    expect(workbook.getNamedExpressionValue("Rate")).toEqual(beforeRateValue);
  });

  it("returns changes in deterministic order for cells and named expressions", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    const changes = workbook.batch(() => {
      workbook.setCellContents(cell(sheetId, 0, 1), 20);
      workbook.setCellContents(cell(sheetId, 0, 0), 10);
      workbook.addNamedExpression("Zulu", "=1");
      workbook.addNamedExpression("Alpha", "=2");
    });

    expect(
      changes.map((change) =>
        change.kind === "cell" ? `${change.kind}:${change.a1}` : `${change.kind}:${change.name}`,
      ),
    ).toEqual(["cell:A1", "cell:B1", "named-expression:Alpha", "named-expression:Zulu"]);
  });

  it("supports once listeners, address formatting, range dependency helpers, and tuple axis operations", () => {
    const workbook = HeadlessWorkbook.buildFromSheets({
      Data: [[1, 2, "=A1+B1"]],
    });
    const sheetId = workbook.getSheetId("Data")!;
    let valuesUpdatedEvents = 0;

    workbook.once("valuesUpdated", () => {
      valuesUpdatedEvents += 1;
    });

    expect(workbook.simpleCellAddressToString(cell(sheetId, 0, 2))).toBe("C1");
    expect(
      workbook.simpleCellAddressToString(cell(sheetId, 0, 2), { includeSheetName: true }),
    ).toBe("Data!C1");

    expect(
      workbook.getCellDependents({ start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 1) }),
    ).toContainEqual({
      kind: "cell",
      address: cell(sheetId, 0, 2),
    });

    workbook.setCellContents(cell(sheetId, 1, 0), 10);
    workbook.setCellContents(cell(sheetId, 1, 1), 20);

    expect(valuesUpdatedEvents).toBe(1);

    workbook.addRows(sheetId, [1, 1]);
    expect(workbook.getSheetDimensions(sheetId).height).toBe(3);

    workbook.swapColumnIndexes(sheetId, [[0, 1]]);
    expect(workbook.getCellSerialized(cell(sheetId, 0, 0))).toBe(2);
    expect(workbook.getCellSerialized(cell(sheetId, 0, 1))).toBe(1);
  });

  it("uses HyperFormula-like optional returns for missing lookups and formula grids", () => {
    const workbook = HeadlessWorkbook.buildFromArray([[1, "=A1+1"]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    expect(workbook.getSheetId("Missing")).toBeUndefined();
    expect(workbook.getSheetName(99)).toBeUndefined();
    expect(workbook.simpleCellAddressFromString("not-an-address")).toBeUndefined();
    expect(workbook.simpleCellRangeFromString("A1")).toBeUndefined();
    expect(workbook.getNamedExpression("Missing")).toBeUndefined();
    expect(workbook.getNamedExpressionFormula("Missing")).toBeUndefined();
    expect(workbook.getNamedExpressionValue("Missing")).toBeUndefined();
    expect(
      workbook.getRangeFormulas({ start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 1) }),
    ).toEqual([[undefined, "=A1+1"]]);
  });

  it("applies function translations to registered languages and exposes license validity", () => {
    HeadlessWorkbook.registerLanguage(TEST_LANGUAGE_CODE, { functions: {} });
    HeadlessWorkbook.registerFunctionPlugin(
      {
        id: "custom-math",
        implementedFunctions: {
          DOUBLE: { method: "DOUBLE" },
        },
      },
      {
        [TEST_LANGUAGE_CODE]: {
          DOUBLE: "DUPLO",
        },
      },
    );

    expect(HeadlessWorkbook.getRegisteredFunctionNames(TEST_LANGUAGE_CODE)).toContain("DUPLO");
    expect(HeadlessWorkbook.buildEmpty().licenseKeyValidityState).toBe("valid");
    expect(HeadlessWorkbook.buildEmpty({ licenseKey: "" }).licenseKeyValidityState).toBe("missing");
  });
});
