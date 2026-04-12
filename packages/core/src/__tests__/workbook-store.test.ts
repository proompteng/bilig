import { describe, expect, it } from "vitest";
import { createCellNumberFormatRecord } from "@bilig/protocol";
import { WorkbookStore } from "../workbook-store.js";

describe("WorkbookStore", () => {
  it("does not mutate existing style ranges when bulk style restoration includes an unknown style", () => {
    const workbook = new WorkbookStore("style-ranges");
    workbook.createSheet("Sheet1");
    workbook.upsertCellStyle({ id: "style-a", font: { bold: true } });
    workbook.setStyleRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      "style-a",
    );

    expect(() =>
      workbook.setStyleRanges("Sheet1", [
        {
          range: { sheetName: "Sheet1", startAddress: "C1", endAddress: "C2" },
          styleId: "style-missing",
        },
      ]),
    ).toThrow("Unknown cell style: style-missing");

    expect(workbook.listStyleRanges("Sheet1")).toEqual([
      {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        styleId: "style-a",
      },
    ]);
  });

  it("does not mutate existing format ranges when bulk format restoration includes an unknown format", () => {
    const workbook = new WorkbookStore("format-ranges");
    workbook.createSheet("Sheet1");
    workbook.upsertCellNumberFormat(createCellNumberFormatRecord("format-money", "$0.00"));
    workbook.setFormatRange(
      { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
      "format-money",
    );

    expect(() =>
      workbook.setFormatRanges("Sheet1", [
        {
          range: { sheetName: "Sheet1", startAddress: "C1", endAddress: "C2" },
          formatId: "format-missing",
        },
      ]),
    ).toThrow("Unknown cell number format: format-missing");

    expect(workbook.listFormatRanges("Sheet1")).toEqual([
      {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B2" },
        formatId: "format-money",
      },
    ]);
  });

  it("normalizes filter and sort ranges so equivalent reversed bounds reuse the same record", () => {
    const workbook = new WorkbookStore("normalized-ranges");
    workbook.createSheet("Sheet1");
    const reversedRange = {
      sheetName: "Sheet1",
      startAddress: "C3",
      endAddress: "A1",
    } as const;
    const normalizedRange = {
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "C3",
    } as const;

    workbook.setFilter("Sheet1", reversedRange);
    workbook.setFilter("Sheet1", normalizedRange);
    workbook.setSort("Sheet1", reversedRange, [{ keyAddress: "B1", direction: "asc" }]);
    workbook.setSort("Sheet1", normalizedRange, [{ keyAddress: "B1", direction: "desc" }]);

    expect(workbook.listFilters("Sheet1")).toEqual([
      { sheetName: "Sheet1", range: normalizedRange },
    ]);
    expect(workbook.getFilter("Sheet1", reversedRange)).toEqual({
      sheetName: "Sheet1",
      range: normalizedRange,
    });
    expect(workbook.deleteFilter("Sheet1", reversedRange)).toBe(true);
    expect(workbook.listFilters("Sheet1")).toEqual([]);

    expect(workbook.listSorts("Sheet1")).toEqual([
      {
        sheetName: "Sheet1",
        range: normalizedRange,
        keys: [{ keyAddress: "B1", direction: "desc" }],
      },
    ]);
    expect(workbook.getSort("Sheet1", reversedRange)).toEqual({
      sheetName: "Sheet1",
      range: normalizedRange,
      keys: [{ keyAddress: "B1", direction: "desc" }],
    });
    expect(workbook.deleteSort("Sheet1", reversedRange)).toBe(true);
    expect(workbook.listSorts("Sheet1")).toEqual([]);
  });

  it("normalizes data validation ranges so equivalent reversed bounds reuse the same record", () => {
    const workbook = new WorkbookStore("normalized-data-validations");
    workbook.createSheet("Sheet1");
    const reversedRange = {
      sheetName: "Sheet1",
      startAddress: "C3",
      endAddress: "A1",
    } as const;
    const normalizedRange = {
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "C3",
    } as const;

    workbook.setDataValidation({
      range: reversedRange,
      rule: {
        kind: "list",
        values: ["Draft", "Final"],
      },
      allowBlank: false,
    });
    workbook.setDataValidation({
      range: normalizedRange,
      rule: {
        kind: "list",
        values: ["Live", "Archived"],
      },
      allowBlank: true,
    });

    expect(workbook.listDataValidations("Sheet1")).toEqual([
      {
        range: normalizedRange,
        rule: {
          kind: "list",
          values: ["Live", "Archived"],
        },
        allowBlank: true,
      },
    ]);
    expect(workbook.getDataValidation("Sheet1", reversedRange)).toEqual({
      range: normalizedRange,
      rule: {
        kind: "list",
        values: ["Live", "Archived"],
      },
      allowBlank: true,
    });
    expect(workbook.deleteDataValidation("Sheet1", reversedRange)).toBe(true);
    expect(workbook.listDataValidations("Sheet1")).toEqual([]);
  });

  it("normalizes spill and pivot addresses so case-only variants reuse the same record", () => {
    const workbook = new WorkbookStore("normalized-addresses");
    workbook.createSheet("Sheet1");

    workbook.setSpill("Sheet1", "b2", 2, 3);
    workbook.setSpill("Sheet1", "B2", 4, 1);

    expect(workbook.listSpills()).toEqual([
      { sheetName: "Sheet1", address: "B2", rows: 4, cols: 1 },
    ]);
    expect(workbook.getSpill("Sheet1", "b2")).toEqual({
      sheetName: "Sheet1",
      address: "B2",
      rows: 4,
      cols: 1,
    });
    expect(workbook.deleteSpill("Sheet1", "b2")).toBe(true);
    expect(workbook.listSpills()).toEqual([]);

    workbook.setPivot({
      name: " RevenuePivot ",
      sheetName: "Sheet1",
      address: "c3",
      source: { sheetName: "Data", startAddress: "a1", endAddress: "b4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });
    workbook.setPivot({
      name: "RevenuePivot",
      sheetName: "Sheet1",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "count" }],
      rows: 4,
      cols: 2,
    });

    expect(workbook.listPivots()).toEqual([
      {
        name: "RevenuePivot",
        sheetName: "Sheet1",
        address: "C3",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
        groupBy: ["Region"],
        values: [{ field: "Sales", summarizeBy: "count" }],
        rows: 4,
        cols: 2,
      },
    ]);
    expect(workbook.getPivot("Sheet1", "c3")).toEqual({
      name: "RevenuePivot",
      sheetName: "Sheet1",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "count" }],
      rows: 4,
      cols: 2,
    });
    expect(workbook.deletePivot("Sheet1", "c3")).toBe(true);
    expect(workbook.listPivots()).toEqual([]);
  });

  it("renames sheet-scoped metadata through the store without leaving stale keys behind", () => {
    const workbook = new WorkbookStore("rename-metadata");
    workbook.createSheet("Source");
    workbook.setFreezePane("Source", 1, 2);
    workbook.setFilter("Source", {
      sheetName: "Source",
      startAddress: "C3",
      endAddress: "A1",
    });
    workbook.setSpill("Source", "b2", 2, 3);
    workbook.setPivot({
      name: "RevenuePivot",
      sheetName: "Source",
      address: "c3",
      source: { sheetName: "Source", startAddress: "b4", endAddress: "a1" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });

    expect(workbook.renameSheet("Source", "Renamed")?.name).toBe("Renamed");

    expect(workbook.getFreezePane("Source")).toBeUndefined();
    expect(workbook.getFreezePane("Renamed")).toEqual({
      sheetName: "Renamed",
      rows: 1,
      cols: 2,
    });
    expect(
      workbook.getFilter("Renamed", {
        sheetName: "Renamed",
        startAddress: "A1",
        endAddress: "C3",
      }),
    ).toEqual({
      sheetName: "Renamed",
      range: { sheetName: "Renamed", startAddress: "A1", endAddress: "C3" },
    });
    expect(workbook.getSpill("Renamed", "B2")).toEqual({
      sheetName: "Renamed",
      address: "B2",
      rows: 2,
      cols: 3,
    });
    expect(workbook.getPivot("Renamed", "C3")).toEqual({
      name: "RevenuePivot",
      sheetName: "Renamed",
      address: "C3",
      source: { sheetName: "Renamed", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });
  });

  it("restores sparse column axis entries without synthesizing blank identities", () => {
    const workbook = new WorkbookStore("sparse-axis-restore");
    workbook.createSheet("Sheet1");
    workbook.setColumnMetadata("Sheet1", 1, 1, 120, false);

    const captured = workbook.snapshotColumnAxisEntries("Sheet1", 0, 2);
    expect(captured).toEqual([{ id: "column-1", index: 1, size: 120, hidden: false }]);

    workbook.deleteColumns("Sheet1", 0, 2);
    workbook.insertColumns("Sheet1", 0, 2, captured);

    expect(workbook.listColumnAxisEntries("Sheet1")).toEqual([
      { id: "column-1", index: 1, size: 120, hidden: false },
    ]);
  });
});
