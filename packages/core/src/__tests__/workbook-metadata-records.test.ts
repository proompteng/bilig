import { describe, expect, it } from "vitest";
import {
  cloneChartRecord,
  cloneCommentEntryRecord,
  cloneCommentThreadRecord,
  cloneConditionalFormatRecord,
  cloneConditionalFormatRule,
  cloneDataValidationRecord,
  cloneDataValidationRule,
  cloneDefinedNameRecord,
  cloneDefinedNameValue,
  cloneFilterRecord,
  cloneImageRecord,
  cloneNoteRecord,
  clonePivotRecord,
  clonePropertyRecord,
  cloneRangeProtectionRecord,
  cloneShapeRecord,
  cloneSheetProtectionRecord,
  cloneSortKeyRecord,
  cloneSortRecord,
  cloneSpillRecord,
  cloneTableRecord,
  commentThreadKey,
  conditionalFormatKey,
  dataValidationKey,
  deleteRecordsBySheet,
  filterKey,
  noteKey,
  rangeProtectionKey,
  rekeyRecords,
  sortKey,
  spillKey,
  tableKey,
} from "../workbook-metadata-records.js";

describe("workbook metadata records", () => {
  it("clones defined names, properties, tables, filters, and sorts without leaking mutation", () => {
    const definedName = {
      name: "SalesRange",
      value: {
        kind: "range-ref" as const,
        sheetName: "Data",
        startAddress: "B4",
        endAddress: "A1",
      },
    };
    const property = { key: "Author", value: "greg" };
    const table = {
      name: "Revenue",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "C10",
      columnNames: ["Region", "Sales"],
      headerRow: true,
      totalsRow: false,
    };
    const filter = {
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "C3", endAddress: "A1" },
    };
    const sort = {
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "D5", endAddress: "B2" },
      keys: [{ keyAddress: "B1", direction: "asc" as const }],
    };

    const clonedDefinedName = cloneDefinedNameRecord(definedName);
    const clonedProperty = clonePropertyRecord(property);
    const clonedTable = cloneTableRecord(table);
    const clonedFilter = cloneFilterRecord(filter);
    const clonedSort = cloneSortRecord(sort);
    const firstSortKey = sort.keys[0];
    if (!firstSortKey) {
      throw new Error("Expected a sort key");
    }
    const clonedSortKey = cloneSortKeyRecord(firstSortKey);

    definedName.value.startAddress = "Z9";
    table.columnNames[0] = "Changed";
    filter.range.startAddress = "Z9";
    firstSortKey.keyAddress = "Z9";

    expect(cloneDefinedNameValue(true)).toBe(true);
    expect(cloneDefinedNameValue("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
    expect(cloneDefinedNameValue({ kind: "scalar", value: 7 })).toEqual({
      kind: "scalar",
      value: 7,
    });
    expect(cloneDefinedNameValue({ kind: "cell-ref", sheetName: "Sheet1", address: "A1" })).toEqual(
      {
        kind: "cell-ref",
        sheetName: "Sheet1",
        address: "A1",
      },
    );
    expect(
      cloneDefinedNameValue({ kind: "structured-ref", tableName: "Revenue", columnName: "Sales" }),
    ).toEqual({
      kind: "structured-ref",
      tableName: "Revenue",
      columnName: "Sales",
    });
    expect(cloneDefinedNameValue({ kind: "formula", formula: "=SUM(A1:A3)" })).toEqual({
      kind: "formula",
      formula: "=SUM(A1:A3)",
    });

    expect(clonedDefinedName).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName: "Data",
        startAddress: "B4",
        endAddress: "A1",
      },
    });
    expect(clonedProperty).toEqual(property);
    expect(clonedTable).toEqual({
      ...table,
      columnNames: ["Region", "Sales"],
    });
    expect(clonedFilter).toEqual({
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "C3", endAddress: "A1" },
    });
    expect(clonedSort).toEqual({
      sheetName: "Sheet1",
      range: { sheetName: "Sheet1", startAddress: "D5", endAddress: "B2" },
      keys: [{ keyAddress: "B1", direction: "asc" }],
    });
    expect(clonedSortKey).toEqual({ keyAddress: "B1", direction: "asc" });
  });

  it("clones validation, conditional format, comment, note, protection, pivot, and media records", () => {
    expect(
      cloneDataValidationRule({
        kind: "list",
        values: ["Draft", "Final"],
        source: {
          kind: "range-ref",
          sheetName: "Data",
          startAddress: "B4",
          endAddress: "A1",
        },
      }),
    ).toEqual({
      kind: "list",
      values: ["Draft", "Final"],
      source: {
        kind: "range-ref",
        sheetName: "Data",
        startAddress: "A1",
        endAddress: "B4",
      },
    });
    expect(
      cloneDataValidationRule({
        kind: "list",
        source: { kind: "cell-ref", sheetName: "Data", address: "c3" },
      }),
    ).toEqual({
      kind: "list",
      source: { kind: "cell-ref", sheetName: "Data", address: "C3" },
    });
    expect(
      cloneDataValidationRule({
        kind: "list",
        source: { kind: "named-range", name: "StatusValues" },
      }),
    ).toEqual({
      kind: "list",
      source: { kind: "named-range", name: "StatusValues" },
    });
    expect(
      cloneDataValidationRule({
        kind: "list",
        source: { kind: "structured-ref", tableName: "Revenue", columnName: "Status" },
      }),
    ).toEqual({
      kind: "list",
      source: { kind: "structured-ref", tableName: "Revenue", columnName: "Status" },
    });
    expect(
      cloneDataValidationRule({
        kind: "checkbox",
        checkedValue: "yes",
        uncheckedValue: "no",
      }),
    ).toEqual({
      kind: "checkbox",
      checkedValue: "yes",
      uncheckedValue: "no",
    });
    expect(
      cloneDataValidationRule({
        kind: "whole",
        operator: "between",
        values: [1, 10],
      }),
    ).toEqual({
      kind: "whole",
      operator: "between",
      values: [1, 10],
    });
    expect(
      cloneDataValidationRule({
        kind: "decimal",
        operator: "greaterThan",
        values: [0],
      }),
    ).toEqual({
      kind: "decimal",
      operator: "greaterThan",
      values: [0],
    });
    expect(
      cloneDataValidationRule({
        kind: "date",
        operator: "equal",
        values: [45000],
      }),
    ).toEqual({
      kind: "date",
      operator: "equal",
      values: [45000],
    });
    expect(
      cloneDataValidationRule({
        kind: "time",
        operator: "lessThan",
        values: [0.5],
      }),
    ).toEqual({
      kind: "time",
      operator: "lessThan",
      values: [0.5],
    });
    expect(
      cloneDataValidationRule({
        kind: "textLength",
        operator: "notBetween",
        values: [3, 8],
      }),
    ).toEqual({
      kind: "textLength",
      operator: "notBetween",
      values: [3, 8],
    });

    expect(
      cloneDataValidationRecord({
        range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "A1" },
        rule: { kind: "list", values: ["Draft"] },
        allowBlank: false,
        showDropdown: true,
        promptTitle: "Status",
        promptMessage: "Pick one",
        errorStyle: "stop",
        errorTitle: "Invalid",
        errorMessage: "Nope",
      }),
    ).toEqual({
      range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "A1" },
      rule: { kind: "list", values: ["Draft"] },
      allowBlank: false,
      showDropdown: true,
      promptTitle: "Status",
      promptMessage: "Pick one",
      errorStyle: "stop",
      errorTitle: "Invalid",
      errorMessage: "Nope",
    });

    expect(
      cloneConditionalFormatRule({
        kind: "cellIs",
        operator: "greaterThan",
        values: [10],
      }),
    ).toEqual({
      kind: "cellIs",
      operator: "greaterThan",
      values: [10],
    });
    expect(
      cloneConditionalFormatRule({
        kind: "textContains",
        text: "pear",
        caseSensitive: true,
      }),
    ).toEqual({
      kind: "textContains",
      text: "pear",
      caseSensitive: true,
    });
    expect(cloneConditionalFormatRule({ kind: "formula", formula: "=A1>0" })).toEqual({
      kind: "formula",
      formula: "=A1>0",
    });
    expect(cloneConditionalFormatRule({ kind: "blanks" })).toEqual({ kind: "blanks" });
    expect(cloneConditionalFormatRule({ kind: "notBlanks" })).toEqual({ kind: "notBlanks" });
    expect(
      cloneConditionalFormatRecord({
        id: "cf-1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        rule: { kind: "formula", formula: "=A1>0" },
        style: { fill: { backgroundColor: "#ff0" } },
        stopIfTrue: true,
        priority: 1,
      }),
    ).toEqual({
      id: "cf-1",
      range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      rule: { kind: "formula", formula: "=A1>0" },
      style: { fill: { backgroundColor: "#ff0" } },
      stopIfTrue: true,
      priority: 1,
    });

    expect(cloneSheetProtectionRecord({ sheetName: "Sheet1", hideFormulas: true })).toEqual({
      sheetName: "Sheet1",
      hideFormulas: true,
    });
    expect(
      cloneRangeProtectionRecord({
        id: "protect-1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        hideFormulas: true,
      }),
    ).toEqual({
      id: "protect-1",
      range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      hideFormulas: true,
    });
    expect(
      cloneCommentEntryRecord({
        id: "comment-1",
        body: "Check this",
        authorUserId: "user-1",
        authorDisplayName: "Greg",
        createdAtUnixMs: 123,
      }),
    ).toEqual({
      id: "comment-1",
      body: "Check this",
      authorUserId: "user-1",
      authorDisplayName: "Greg",
      createdAtUnixMs: 123,
    });
    expect(
      cloneCommentThreadRecord({
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "c3",
        comments: [{ id: "comment-1", body: "Check this" }],
        resolved: true,
        resolvedByUserId: "user-1",
        resolvedAtUnixMs: 456,
      }),
    ).toEqual({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "C3",
      comments: [{ id: "comment-1", body: "Check this" }],
      resolved: true,
      resolvedByUserId: "user-1",
      resolvedAtUnixMs: 456,
    });
    expect(cloneNoteRecord({ sheetName: "Sheet1", address: "d4", text: "Note" })).toEqual({
      sheetName: "Sheet1",
      address: "D4",
      text: "Note",
    });
    expect(
      clonePivotRecord({
        name: "RevenuePivot",
        sheetName: "Sheet1",
        address: "C3",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
        groupBy: ["Region"],
        values: [{ field: "Sales", summarizeBy: "sum" }],
        rows: 3,
        cols: 2,
      }),
    ).toEqual({
      name: "RevenuePivot",
      sheetName: "Sheet1",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });
    expect(
      cloneChartRecord({
        id: "chart-1",
        sheetName: "Sheet1",
        address: "C4",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
        chartType: "column",
        seriesOrientation: "rows",
        firstRowAsHeaders: true,
        firstColumnAsLabels: false,
        title: "Revenue",
        legendPosition: "bottom",
        rows: 5,
        cols: 6,
      }),
    ).toEqual({
      id: "chart-1",
      sheetName: "Sheet1",
      address: "C4",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      chartType: "column",
      seriesOrientation: "rows",
      firstRowAsHeaders: true,
      firstColumnAsLabels: false,
      title: "Revenue",
      legendPosition: "bottom",
      rows: 5,
      cols: 6,
    });
    expect(
      cloneImageRecord({
        id: "image-1",
        sheetName: "Sheet1",
        address: "d6",
        sourceUrl: "https://example.com/image.png",
        rows: 3,
        cols: 4,
        altText: "Revenue image",
      }),
    ).toEqual({
      id: "image-1",
      sheetName: "Sheet1",
      address: "D6",
      sourceUrl: "https://example.com/image.png",
      rows: 3,
      cols: 4,
      altText: "Revenue image",
    });
    expect(
      cloneShapeRecord({
        id: "shape-1",
        sheetName: "Sheet1",
        address: "e7",
        shapeType: "textBox",
        rows: 2,
        cols: 3,
        text: "Quarterly note",
        fillColor: "#ffeeaa",
        strokeColor: "#222222",
      }),
    ).toEqual({
      id: "shape-1",
      sheetName: "Sheet1",
      address: "E7",
      shapeType: "textBox",
      rows: 2,
      cols: 3,
      text: "Quarterly note",
      fillColor: "#ffeeaa",
      strokeColor: "#222222",
    });
    expect(cloneSpillRecord({ sheetName: "Sheet1", address: "A1", rows: 2, cols: 3 })).toEqual({
      sheetName: "Sheet1",
      address: "A1",
      rows: 2,
      cols: 3,
    });
  });

  it("builds canonical metadata keys and validates required ids", () => {
    expect(filterKey("Sheet1", { sheetName: "Sheet1", startAddress: "C3", endAddress: "A1" })).toBe(
      "Sheet1:A1:C3",
    );
    expect(sortKey("Sheet1", { sheetName: "Sheet1", startAddress: "D4", endAddress: "B2" })).toBe(
      "Sheet1:B2:D4",
    );
    expect(
      dataValidationKey("Sheet1", { sheetName: "Sheet1", startAddress: "D4", endAddress: "B2" }),
    ).toBe("Sheet1:B2:D4");
    expect(commentThreadKey("Sheet1", "c3")).toBe("Sheet1!C3");
    expect(noteKey("Sheet1", "d4")).toBe("Sheet1!D4");
    expect(tableKey(" revenue ")).toBe("REVENUE");
    expect(spillKey("Sheet1", "e5")).toBe("Sheet1!E5");
    expect(conditionalFormatKey(" cf-1 ")).toBe("cf-1");
    expect(rangeProtectionKey(" protect-1 ")).toBe("protect-1");
    expect(() => conditionalFormatKey("   ")).toThrow("Conditional format id must be non-empty");
    expect(() => rangeProtectionKey("   ")).toThrow("Range protection id must be non-empty");
  });

  it("deletes and rekeys workbook metadata records across record families", () => {
    const notes = new Map<string, { sheetName: string; address: string; text: string }>([
      ["Sheet1!A1", { sheetName: "Sheet1", address: "A1", text: "keep" }],
      ["Sheet2!B2", { sheetName: "Sheet2", address: "B2", text: "drop" }],
    ]);
    deleteRecordsBySheet(notes, "Sheet2", (record) => record.sheetName);
    expect([...notes.keys()]).toEqual(["Sheet1!A1"]);

    const expectSingleRekey = <T extends object>(
      bucket: Map<string, T>,
      rewrite: (record: T) => T,
      expectedKey: string,
    ) => {
      rekeyRecords(bucket, rewrite);
      expect([...bucket.keys()]).toEqual([expectedKey]);
    };

    expectSingleRekey(
      new Map([["old", { sheetName: "Sheet1", rows: 1, cols: 2 }]]),
      (record) => ({ ...record, sheetName: "Renamed" }),
      "Renamed",
    );
    expectSingleRekey(
      new Map([["old", { sheetName: "Sheet1", start: 0, count: 2, size: 24, hidden: null }]]),
      (record) => ({ ...record, sheetName: "Renamed", start: 3 }),
      "Renamed:3:2",
    );
    expectSingleRekey(
      new Map([["old", { sheetName: "Sheet1", hideFormulas: true }]]),
      (record) => ({ ...record, sheetName: "Renamed" }),
      "Renamed",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            sheetName: "Sheet1",
            range: { sheetName: "Sheet1", startAddress: "C3", endAddress: "A1" },
          },
        ],
      ]),
      (record) => ({ ...record, sheetName: "Renamed" }),
      "Renamed:A1:C3",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            sheetName: "Sheet1",
            range: { sheetName: "Sheet1", startAddress: "D4", endAddress: "B2" },
            keys: [{ keyAddress: "B1", direction: "asc" as const }],
          },
        ],
      ]),
      (record) => ({ ...record, sheetName: "Renamed" }),
      "Renamed:B2:D4",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            range: { sheetName: "Sheet1", startAddress: "D4", endAddress: "B2" },
            rule: { kind: "list" as const, values: ["Draft"] },
          },
        ],
      ]),
      (record) => ({
        ...record,
        range: { ...record.range, sheetName: "Renamed" },
      }),
      "Renamed:B2:D4",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            id: " cf-1 ",
            range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
            rule: { kind: "blanks" as const },
            style: {},
          },
        ],
      ]),
      (record) => ({ ...record, id: "cf-2" }),
      "cf-2",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            id: " protect-1 ",
            range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
          },
        ],
      ]),
      (record) => ({ ...record, id: "protect-2" }),
      "protect-2",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            threadId: "thread-1",
            sheetName: "Sheet1",
            address: "c3",
            comments: [{ id: "comment-1", body: "check" }],
          },
        ],
      ]),
      (record) => ({ ...record, address: "d4" }),
      "Sheet1!D4",
    );
    expectSingleRekey(
      new Map([["old", { sheetName: "Sheet1", address: "d4", text: "note" }]]),
      (record) => ({ ...record, address: "e5" }),
      "Sheet1!E5",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            name: " Revenue ",
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B4",
            columnNames: ["Region", "Sales"],
            headerRow: true,
            totalsRow: false,
          },
        ],
      ]),
      (record) => ({ ...record, name: "Pipeline" }),
      "PIPELINE",
    );
    expectSingleRekey(
      new Map([["old", { sheetName: "Sheet1", address: "a1", rows: 2, cols: 3 }]]),
      (record) => ({ ...record, address: "b2" }),
      "Sheet1!B2",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            name: "RevenuePivot",
            sheetName: "Sheet1",
            address: "c3",
            source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
            groupBy: ["Region"],
            values: [{ field: "Sales", summarizeBy: "sum" as const }],
            rows: 3,
            cols: 2,
          },
        ],
      ]),
      (record) => ({ ...record, address: "d4" }),
      "Sheet1!D4",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            id: " chart-1 ",
            sheetName: "Sheet1",
            address: "c4",
            source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
            chartType: "column" as const,
            rows: 4,
            cols: 5,
          },
        ],
      ]),
      (record) => ({ ...record, id: "chart-2" }),
      "CHART-2",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            id: " image-1 ",
            sheetName: "Sheet1",
            address: "d6",
            sourceUrl: "https://example.com/image.png",
            rows: 3,
            cols: 4,
          },
        ],
      ]),
      (record) => ({ ...record, id: "image-2" }),
      "IMAGE-2",
    );
    expectSingleRekey(
      new Map([
        [
          "old",
          {
            id: " shape-1 ",
            sheetName: "Sheet1",
            address: "e7",
            shapeType: "textBox" as const,
            rows: 2,
            cols: 3,
          },
        ],
      ]),
      (record) => ({ ...record, id: "shape-2" }),
      "SHAPE-2",
    );
  });
});
