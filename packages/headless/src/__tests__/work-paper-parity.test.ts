import { afterEach, describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";

import {
  WorkPaperExpectedOneOfValuesError,
  WorkPaperExpectedValueOfTypeError,
  WorkPaperLanguageAlreadyRegisteredError,
  WorkPaperLanguageNotRegisteredError,
  WorkPaperNamedExpressionDoesNotExistError,
  WorkPaperNamedExpressionNameIsAlreadyTakenError,
  WorkPaperNoOperationToRedoError,
  WorkPaperNoOperationToUndoError,
  WorkPaperNoRelativeAddressesAllowedError,
  WorkPaperNotAFormulaError,
  WorkPaperNothingToPasteError,
  WorkPaper,
  WorkPaperCellAddress,
  WorkPaperConfig,
  WorkPaperFunctionPluginDefinition,
} from "../index.js";

const TEST_LANGUAGE_CODE = "hf-parity";

const CUSTOM_PLUGIN: WorkPaperFunctionPluginDefinition = {
  id: "hf-parity-plugin",
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
};

function createConfigFixture(): WorkPaperConfig {
  return {
    accentSensitive: true,
    caseSensitive: true,
    caseFirst: "upper",
    chooseAddressMappingPolicy: { mode: "dense" },
    context: { requestId: "ctx-1" },
    currencySymbol: ["$", "USD"],
    dateFormats: ["YYYY-MM-DD"],
    functionArgSeparator: ";",
    decimalSeparator: ",",
    evaluateNullToZero: false,
    functionPlugins: [CUSTOM_PLUGIN],
    ignorePunctuation: true,
    language: TEST_LANGUAGE_CODE,
    ignoreWhiteSpace: "any",
    leapYear1900: false,
    licenseKey: "",
    localeLang: "pl-PL",
    matchWholeCell: false,
    arrayColumnSeparator: ";",
    arrayRowSeparator: "|",
    maxRows: 2048,
    maxColumns: 256,
    nullDate: { year: 1904, month: 1, day: 1 },
    nullYear: 50,
    parseDateTime: () => ({
      year: 2024,
      month: 1,
      day: 2,
      hours: 3,
      minutes: 4,
      seconds: 5,
    }),
    precisionEpsilon: 1e-8,
    precisionRounding: 10,
    stringifyDateTime: () => "date",
    stringifyDuration: () => "duration",
    smartRounding: false,
    thousandSeparator: ".",
    timeFormats: ["HH:mm:ss"],
    useArrayArithmetic: false,
    useColumnIndex: true,
    useStats: false,
    undoLimit: 25,
    useRegularExpressions: false,
    useWildcards: false,
  };
}

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col };
}

afterEach(() => {
  WorkPaper.unregisterAllFunctions();
  if (WorkPaper.getRegisteredLanguagesCodes().includes(TEST_LANGUAGE_CODE)) {
    WorkPaper.unregisterLanguage(TEST_LANGUAGE_CODE);
  }
});

describe("WorkPaper parity surface", () => {
  it("covers static factories, config inventory, and registry APIs", () => {
    expect(Object.keys(WorkPaper.defaultConfig).toSorted()).toEqual([
      "accentSensitive",
      "arrayColumnSeparator",
      "arrayRowSeparator",
      "caseFirst",
      "caseSensitive",
      "chooseAddressMappingPolicy",
      "context",
      "currencySymbol",
      "dateFormats",
      "decimalSeparator",
      "evaluateNullToZero",
      "functionArgSeparator",
      "functionPlugins",
      "ignorePunctuation",
      "ignoreWhiteSpace",
      "language",
      "leapYear1900",
      "licenseKey",
      "localeLang",
      "matchWholeCell",
      "maxColumns",
      "maxRows",
      "nullDate",
      "nullYear",
      "parseDateTime",
      "precisionEpsilon",
      "precisionRounding",
      "smartRounding",
      "stringifyDateTime",
      "stringifyDuration",
      "thousandSeparator",
      "timeFormats",
      "undoLimit",
      "useArrayArithmetic",
      "useColumnIndex",
      "useRegularExpressions",
      "useStats",
      "useWildcards",
    ]);

    expect(WorkPaper.buildEmpty().getSheetNames()).toEqual([]);
    expect(WorkPaper.buildFromArray([[1]]).getSheetNames()).toEqual(["Sheet1"]);
    expect(
      WorkPaper.buildFromSheets({
        Alpha: [[1]],
        Beta: [[2]],
      }).getSheetNames(),
    ).toEqual(["Alpha", "Beta"]);

    expect(() => WorkPaper.getLanguage(TEST_LANGUAGE_CODE)).toThrow(
      WorkPaperLanguageNotRegisteredError,
    );

    WorkPaper.registerLanguage(TEST_LANGUAGE_CODE, { functions: {} });
    expect(WorkPaper.getRegisteredLanguagesCodes()).toContain(TEST_LANGUAGE_CODE);
    expect(WorkPaper.getLanguage(TEST_LANGUAGE_CODE)).toEqual({ functions: {} });
    expect(() => WorkPaper.registerLanguage(TEST_LANGUAGE_CODE, { functions: {} })).toThrow(
      WorkPaperLanguageAlreadyRegisteredError,
    );

    WorkPaper.registerFunctionPlugin(CUSTOM_PLUGIN, {
      [TEST_LANGUAGE_CODE]: { DOUBLE: "DUPLO" },
    });

    expect(WorkPaper.getRegisteredFunctionNames(TEST_LANGUAGE_CODE)).toContain("DUPLO");
    expect(WorkPaper.getFunctionPlugin("DOUBLE")?.id).toBe(CUSTOM_PLUGIN.id);
    expect(WorkPaper.getAllFunctionPlugins().map((plugin) => plugin.id)).toContain(
      CUSTOM_PLUGIN.id,
    );
  });

  it("covers all config keys through build, clone, and update paths", () => {
    WorkPaper.registerLanguage(TEST_LANGUAGE_CODE, { functions: {} });
    WorkPaper.registerFunctionPlugin(CUSTOM_PLUGIN);

    const config = createConfigFixture();
    const workbook = WorkPaper.buildFromArray([[1]], config);
    const snapshot = workbook.getConfig();

    expect(snapshot).toMatchObject({
      accentSensitive: true,
      caseSensitive: true,
      caseFirst: "upper",
      currencySymbol: ["$", "USD"],
      dateFormats: ["YYYY-MM-DD"],
      functionArgSeparator: ";",
      decimalSeparator: ",",
      evaluateNullToZero: false,
      ignorePunctuation: true,
      language: TEST_LANGUAGE_CODE,
      ignoreWhiteSpace: "any",
      leapYear1900: false,
      licenseKey: "",
      localeLang: "pl-PL",
      matchWholeCell: false,
      arrayColumnSeparator: ";",
      arrayRowSeparator: "|",
      maxRows: 2048,
      maxColumns: 256,
      nullDate: { year: 1904, month: 1, day: 1 },
      nullYear: 50,
      precisionEpsilon: 1e-8,
      precisionRounding: 10,
      smartRounding: false,
      thousandSeparator: ".",
      timeFormats: ["HH:mm:ss"],
      useArrayArithmetic: false,
      useColumnIndex: true,
      useStats: false,
      undoLimit: 25,
      useRegularExpressions: false,
      useWildcards: false,
    });
    expect(snapshot.chooseAddressMappingPolicy).toEqual({ mode: "dense" });
    expect(snapshot.context).toEqual({ requestId: "ctx-1" });
    expect(snapshot.functionPlugins?.map((plugin) => plugin.id)).toEqual([CUSTOM_PLUGIN.id]);
    expect(snapshot.parseDateTime?.("value")).toEqual({
      year: 2024,
      month: 1,
      day: 2,
      hours: 3,
      minutes: 4,
      seconds: 5,
    });
    expect(
      snapshot.stringifyDateTime?.(
        { year: 2024, month: 1, day: 2, hours: 3, minutes: 4, seconds: 5 },
        "YYYY-MM-DD HH:mm:ss",
      ),
    ).toBe("date");
    expect(snapshot.stringifyDuration?.({ hours: 3, minutes: 4, seconds: 5 }, "HH:mm:ss")).toBe(
      "duration",
    );
    expect(workbook.licenseKeyValidityState).toBe("missing");

    snapshot.currencySymbol?.push("MUTATED");
    snapshot.dateFormats?.push("MUTATED");
    snapshot.timeFormats?.push("MUTATED");
    snapshot.functionPlugins?.push({
      id: "mutated",
      implementedFunctions: {},
    });
    if (snapshot.chooseAddressMappingPolicy) {
      snapshot.chooseAddressMappingPolicy.mode = "sparse";
    }
    if (
      snapshot.context &&
      typeof snapshot.context === "object" &&
      !Array.isArray(snapshot.context)
    ) {
      snapshot.context.requestId = "ctx-mutated";
    }

    expect(workbook.getConfig().currencySymbol).toEqual(["$", "USD"]);
    expect(workbook.getConfig().dateFormats).toEqual(["YYYY-MM-DD"]);
    expect(workbook.getConfig().timeFormats).toEqual(["HH:mm:ss"]);
    expect(workbook.getConfig().functionPlugins?.map((plugin) => plugin.id)).toEqual([
      CUSTOM_PLUGIN.id,
    ]);
    expect(workbook.getConfig().chooseAddressMappingPolicy).toEqual({ mode: "dense" });
    expect(workbook.getConfig().context).toEqual({ requestId: "ctx-1" });

    workbook.updateConfig({
      accentSensitive: false,
      caseSensitive: false,
      caseFirst: "lower",
      chooseAddressMappingPolicy: { mode: "sparse" },
      context: { requestId: "ctx-2" },
      currencySymbol: ["EUR"],
      dateFormats: ["DD/MM/YYYY"],
      functionArgSeparator: ",",
      decimalSeparator: ".",
      evaluateNullToZero: true,
      functionPlugins: [],
      ignorePunctuation: false,
      ignoreWhiteSpace: "standard",
      leapYear1900: true,
      licenseKey: "internal",
      localeLang: "en-US",
      matchWholeCell: true,
      arrayColumnSeparator: ",",
      arrayRowSeparator: ";",
      maxRows: 1024,
      maxColumns: 128,
      nullDate: { year: 1899, month: 12, day: 30 },
      nullYear: 30,
      parseDateTime: undefined,
      precisionEpsilon: 1e-13,
      precisionRounding: 14,
      stringifyDateTime: undefined,
      stringifyDuration: undefined,
      smartRounding: true,
      thousandSeparator: ",",
      timeFormats: ["HH:mm"],
      useArrayArithmetic: true,
      useColumnIndex: false,
      useStats: true,
      undoLimit: 100,
      useRegularExpressions: true,
      useWildcards: true,
    });

    expect(workbook.getConfig()).toMatchObject({
      accentSensitive: false,
      caseSensitive: false,
      caseFirst: "lower",
      currencySymbol: ["EUR"],
      dateFormats: ["DD/MM/YYYY"],
      functionArgSeparator: ",",
      decimalSeparator: ".",
      evaluateNullToZero: true,
      functionPlugins: [],
      ignorePunctuation: false,
      ignoreWhiteSpace: "standard",
      leapYear1900: true,
      licenseKey: "internal",
      localeLang: "en-US",
      matchWholeCell: true,
      arrayColumnSeparator: ",",
      arrayRowSeparator: ";",
      maxRows: 1024,
      maxColumns: 128,
      nullDate: { year: 1899, month: 12, day: 30 },
      nullYear: 30,
      precisionEpsilon: 1e-13,
      precisionRounding: 14,
      smartRounding: true,
      thousandSeparator: ",",
      timeFormats: ["HH:mm"],
      useArrayArithmetic: true,
      useColumnIndex: false,
      useStats: true,
      undoLimit: 100,
      useRegularExpressions: true,
      useWildcards: true,
    });
    expect(workbook.getConfig().chooseAddressMappingPolicy).toEqual({ mode: "sparse" });
    expect(workbook.getConfig().context).toEqual({ requestId: "ctx-2" });
    expect(workbook.licenseKeyValidityState).toBe("valid");
  });

  it("rejects invalid typed config hooks and policy values", () => {
    expect(() =>
      // @ts-expect-error intentional invalid runtime input
      WorkPaper.buildEmpty({
        chooseAddressMappingPolicy: { mode: "invalid" },
      }),
    ).toThrow(WorkPaperExpectedOneOfValuesError);
    expect(() =>
      // @ts-expect-error intentional invalid runtime input
      WorkPaper.buildEmpty({
        parseDateTime: 123,
      }),
    ).toThrow(WorkPaperExpectedValueOfTypeError);
    expect(() =>
      // @ts-expect-error intentional invalid runtime input
      WorkPaper.buildEmpty({
        context: { ok: "yes", bad: () => "nope" },
      }),
    ).toThrow(WorkPaperExpectedValueOfTypeError);
  });

  it("covers the read surface, formula helpers, and dependency helpers", () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1, '=HYPERLINK("https://example.com","Docs")', "=A1+1"]],
      Calc: [[10]],
    });
    const dataId = workbook.getSheetId("Data")!;
    const calcId = workbook.getSheetId("Calc")!;

    expect(workbook.getCellValue(cell(dataId, 0, 0))).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(workbook.getCellFormula(cell(dataId, 0, 1))).toBe(
      '=HYPERLINK("https://example.com","Docs")',
    );
    expect(workbook.getCellHyperlink(cell(dataId, 0, 1))).toBe("https://example.com");
    expect(workbook.getCellSerialized(cell(dataId, 0, 2))).toBe("=A1+1");
    expect(
      workbook.getRangeValues({
        start: cell(dataId, 0, 0),
        end: cell(dataId, 0, 2),
      })[0]?.length,
    ).toBe(3);
    expect(
      workbook.getRangeFormulas({
        start: cell(dataId, 0, 0),
        end: cell(dataId, 0, 2),
      }),
    ).toEqual([[undefined, '=HYPERLINK("https://example.com","Docs")', "=A1+1"]]);
    expect(workbook.getSheetValues(dataId)[0]?.length).toBe(3);
    expect(workbook.getSheetFormulas(dataId)[0]?.[2]).toBe("=A1+1");
    expect(workbook.getSheetSerialized(calcId)[0]?.[0]).toBe(10);
    expect(workbook.getAllSheetsValues().Calc?.[0]?.[0]).toMatchObject({
      tag: ValueTag.Number,
      value: 10,
    });
    expect(workbook.getAllSheetsFormulas().Data?.[0]?.[2]).toBe("=A1+1");
    expect(workbook.getAllSheetsSerialized().Data?.[0]?.[1]).toBe(
      '=HYPERLINK("https://example.com","Docs")',
    );
    expect(workbook.getAllSheetsDimensions()).toEqual({
      Calc: { width: 1, height: 1 },
      Data: { width: 3, height: 1 },
    });
    expect(workbook.getSheetDimensions(dataId)).toEqual({ width: 3, height: 1 });
    expect(workbook.getCellDependents(cell(dataId, 0, 0))).toContainEqual({
      kind: "cell",
      address: cell(dataId, 0, 2),
    });
    expect(workbook.getCellPrecedents(cell(dataId, 0, 2))).toContainEqual({
      kind: "cell",
      address: cell(dataId, 0, 0),
    });
    expect(workbook.getCellType(cell(dataId, 0, 0))).toBe("VALUE");
    expect(workbook.doesCellHaveSimpleValue(cell(dataId, 0, 0))).toBe(true);
    expect(workbook.doesCellHaveFormula(cell(dataId, 0, 2))).toBe(true);
    expect(workbook.isCellEmpty(cell(dataId, 1, 0))).toBe(true);
    expect(workbook.isCellPartOfArray(cell(dataId, 0, 2))).toBe(false);
    expect(workbook.getCellValueType(cell(dataId, 0, 2))).toBe("NUMBER");
    expect(workbook.getCellValueDetailedType(cell(dataId, 0, 2))).toBe("NUMBER");
    expect(workbook.getCellValueFormat(cell(dataId, 0, 2))).toBeUndefined();
    expect(workbook.simpleCellAddressFromString("Data!C1")).toEqual(cell(dataId, 0, 2));
    expect(workbook.simpleCellRangeFromString("Data!A1:C1")).toEqual({
      start: cell(dataId, 0, 0),
      end: cell(dataId, 0, 2),
    });
    expect(workbook.simpleCellAddressToString(cell(dataId, 0, 2), dataId)).toBe("C1");
    expect(
      workbook.simpleCellRangeToString(
        {
          start: cell(dataId, 0, 0),
          end: cell(dataId, 0, 2),
        },
        { includeSheetName: true },
      ),
    ).toBe("Data!A1:C1");
    expect(workbook.normalizeFormula("=sum(a1)")).toBe("=SUM(A1)");
    expect(workbook.calculateFormula("=SUM(1,2,3)")).toMatchObject({
      tag: ValueTag.Number,
      value: 6,
    });
    expect(workbook.getNamedExpressionsFromFormula("=Alpha+Beta")).toEqual(["Alpha", "Beta"]);
    expect(workbook.validateFormula("=SUM(1,2)")).toBe(true);
    expect(workbook.validateFormula("SUM(1,2)")).toBe(false);
    expect(workbook.numberToDateTime(2.5)).toEqual({
      year: 1900,
      month: 1,
      day: 2,
      hours: 12,
      minutes: 0,
      seconds: 0,
    });
    expect(workbook.numberToDate(2.5)).toEqual({ year: 1900, month: 1, day: 2 });
    expect(workbook.numberToTime(2.5)).toEqual({ hours: 12, minutes: 0, seconds: 0 });
    expect(() => workbook.normalizeFormula("SUM(1,2)")).toThrow(WorkPaperNotAFormulaError);
  });

  it("covers mutations, preflights, history controls, clipboard, and fill helpers", () => {
    const workbook = WorkPaper.buildFromArray([[1, 2, "=A1+B1"]]);
    const sheetId = workbook.getSheetId("Sheet1")!;

    expect(workbook.isItPossibleToSetCellContents(cell(sheetId, 1, 1), 5)).toBe(true);
    expect(workbook.isItPossibleToAddRows(sheetId, 1, 1)).toBe(true);
    expect(workbook.isItPossibleToAddColumns(sheetId, 1, 1)).toBe(true);
    expect(
      workbook.isItPossibleToMoveCells(
        { start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 0) },
        cell(sheetId, 1, 0),
      ),
    ).toBe(true);

    workbook.setCellContents(cell(sheetId, 1, 0), 10);
    workbook.setCellContents(cell(sheetId, 1, 1), 20);
    workbook.moveCells(
      { start: cell(sheetId, 1, 0), end: cell(sheetId, 1, 1) },
      cell(sheetId, 2, 0),
    );
    workbook.addRows(sheetId, [1, 1]);
    workbook.addColumns(sheetId, [1, 1]);
    workbook.removeColumns(sheetId, [1, 1]);
    workbook.removeRows(sheetId, [1, 1]);
    workbook.moveRows(sheetId, 1, 1, 0);
    workbook.moveColumns(sheetId, 1, 1, 0);
    workbook.setRowOrder(sheetId, [0, 1, 2]);
    workbook.setColumnOrder(sheetId, [0, 1, 2]);

    const fillWorkbook = WorkPaper.buildFromArray([[1, 2, "=A1+B1"]]);
    const fillSheetId = fillWorkbook.getSheetId("Sheet1")!;
    const copied = fillWorkbook.copy({
      start: cell(fillSheetId, 0, 0),
      end: cell(fillSheetId, 0, 2),
    });
    expect(copied[0]?.length).toBe(3);
    expect(fillWorkbook.isClipboardEmpty()).toBe(false);
    expect(
      fillWorkbook.getFillRangeData(
        { start: cell(fillSheetId, 0, 2), end: cell(fillSheetId, 0, 2) },
        { start: cell(fillSheetId, 1, 2), end: cell(fillSheetId, 2, 2) },
      ),
    ).toEqual([["=A2+B2"], ["=A3+B3"]]);
    fillWorkbook.paste(cell(fillSheetId, 3, 0));
    fillWorkbook.cut({
      start: cell(fillSheetId, 3, 0),
      end: cell(fillSheetId, 3, 2),
    });
    fillWorkbook.clearClipboard();
    expect(fillWorkbook.isClipboardEmpty()).toBe(true);
    expect(() => fillWorkbook.paste(cell(fillSheetId, 4, 0))).toThrow(WorkPaperNothingToPasteError);

    workbook.copy({
      start: cell(sheetId, 0, 0),
      end: cell(sheetId, 0, 2),
    });

    expect(workbook.isThereSomethingToUndo()).toBe(true);
    workbook.clearRedoStack();
    workbook.undo();
    expect(workbook.isThereSomethingToRedo()).toBe(true);
    workbook.redo();
    workbook.clearUndoStack();
    expect(workbook.isThereSomethingToUndo()).toBe(false);
    expect(() => workbook.undo()).toThrow(WorkPaperNoOperationToUndoError);
    workbook.clearRedoStack();
    expect(() => workbook.redo()).toThrow(WorkPaperNoOperationToRedoError);
  });

  it("covers sheet lifecycle, named expressions, events, and internal adapters", () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1, "=Rate+1"]],
    });
    const events: string[] = [];
    const detailed: string[] = [];

    workbook.on("sheetAdded", (sheetName) => {
      events.push(`sheetAdded:${sheetName}`);
    });
    workbook.on("sheetRenamed", (oldName, newName) => {
      events.push(`sheetRenamed:${oldName}:${newName}`);
    });
    workbook.on("namedExpressionAdded", (name, changes) => {
      events.push(`nameAdded:${name}:${changes.length}`);
    });
    workbook.on("namedExpressionRemoved", (name, changes) => {
      events.push(`nameRemoved:${name}:${changes.length}`);
    });
    workbook.on("sheetRemoved", (sheetName, changes) => {
      events.push(`sheetRemoved:${sheetName}:${changes.length}`);
    });
    workbook.onDetailed("sheetAdded", (payload) => {
      detailed.push(`sheetId:${payload.sheetId}`);
    });

    const dataId = workbook.getSheetId("Data")!;
    expect(workbook.isItPossibleToAddNamedExpression("Rate", "=1", dataId)).toBe(true);
    workbook.addNamedExpression("Rate", "=1", dataId);
    expect(workbook.getNamedExpressionValue("Rate", dataId)).toMatchObject({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(workbook.getNamedExpressionFormula("Rate", dataId)).toBe("=1");
    expect(workbook.getNamedExpression("Rate", dataId)?.scope).toBe(dataId);
    expect(workbook.listNamedExpressions(dataId)).toEqual(["Rate"]);
    expect(workbook.getAllNamedExpressionsSerialized()).toEqual([
      { name: "Rate", expression: "=1", scope: dataId, options: undefined },
    ]);
    expect(workbook.isItPossibleToChangeNamedExpression("Rate", "=2", dataId)).toBe(true);
    workbook.changeNamedExpression("Rate", "=2", dataId);
    expect(workbook.getNamedExpressionValue("Rate", dataId)).toMatchObject({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(() => workbook.addNamedExpression("Rate", "=3", dataId)).toThrow(
      WorkPaperNamedExpressionNameIsAlreadyTakenError,
    );
    expect(() => workbook.changeNamedExpression("Missing", "=1", dataId)).toThrow(
      WorkPaperNamedExpressionDoesNotExistError,
    );
    expect(() => workbook.addNamedExpression("RelativeName", "=A1", dataId)).toThrow(
      WorkPaperNoRelativeAddressesAllowedError,
    );
    expect(workbook.isItPossibleToRemoveNamedExpression("Rate", dataId)).toBe(true);
    workbook.removeNamedExpression("Rate", dataId);
    expect(workbook.getNamedExpression("Rate", dataId)).toBeUndefined();

    expect(workbook.isItPossibleToAddSheet("Archive")).toBe(true);
    const newSheetName = workbook.addSheet("Archive");
    const archiveId = workbook.getSheetId(newSheetName)!;
    workbook.renameSheet(archiveId, "Archive 2026");
    workbook.setSheetContent(archiveId, [[99]]);
    expect(workbook.getSheetSerialized(archiveId)).toEqual([[99]]);
    expect(workbook.isItPossibleToClearSheet(archiveId)).toBe(true);
    workbook.clearSheet(archiveId);
    expect(workbook.getSheetValues(archiveId)).toEqual([]);
    expect(workbook.isItPossibleToRemoveSheet(archiveId)).toBe(true);
    workbook.removeSheet(archiveId);

    expect(events).toContain("sheetAdded:Archive");
    expect(events).toContain("sheetRenamed:Archive:Archive 2026");
    expect(events.some((event) => event.startsWith("nameAdded:Rate:"))).toBe(true);
    expect(events.some((event) => event.startsWith("nameRemoved:Rate:"))).toBe(true);
    expect(events.some((event) => event.startsWith("sheetRemoved:Archive 2026:"))).toBe(true);
    expect(detailed).toEqual([`sheetId:${archiveId}`]);

    expect(workbook.graph.getDependents(cell(dataId, 0, 0))).toEqual([]);
    expect(
      workbook.rangeMapping.getSerialized({
        start: cell(dataId, 0, 0),
        end: cell(dataId, 0, 1),
      })[0]?.[1],
    ).toBe("=Rate+1");
    expect(workbook.arrayMapping.isPartOfArray(cell(dataId, 0, 1))).toBe(false);
    expect(workbook.sheetMapping.getSheetName(dataId)).toBe("Data");
    expect(workbook.addressMapping.has(cell(dataId, 0, 1))).toBe(true);
    expect(workbook.dependencyGraph.getCellPrecedents(cell(dataId, 0, 1))).toContainEqual({
      kind: "name",
      name: "Rate",
    });
    expect(workbook.evaluator.calculateFormula("=SUM(2,3)")).toMatchObject({
      tag: ValueTag.Number,
      value: 5,
    });
    expect(workbook.columnSearch.find(dataId, 0, (value) => value.tag === ValueTag.Number)).toEqual(
      [cell(dataId, 0, 0)],
    );
    expect(workbook.lazilyTransformingAstService.validateFormula("=A1+1")).toBe(true);
  });
});
