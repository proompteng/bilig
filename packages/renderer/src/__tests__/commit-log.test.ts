import { describe, expect, it } from "vitest";
import { collectDeleteOps, collectMountOps, collectSheetOrderOps, normalizeCommitOps } from "../commit-log.js";
import type { CellDescriptor, SheetDescriptor, WorkbookDescriptor } from "../descriptors.js";

function cell(addr: string, value?: number, formula?: string): CellDescriptor {
  return {
    kind: "Cell",
    props: {
      addr,
      ...(value !== undefined ? { value } : {}),
      ...(formula !== undefined ? { formula } : {})
    },
    parent: null,
    container: null
  };
}

function sheet(name: string, children: CellDescriptor[] = []): SheetDescriptor {
  const descriptor: SheetDescriptor = {
    kind: "Sheet",
    props: { name },
    children,
    parent: null,
    container: null
  };
  children.forEach((child) => {
    child.parent = descriptor;
  });
  return descriptor;
}

function workbook(children: SheetDescriptor[]): WorkbookDescriptor {
  const descriptor: WorkbookDescriptor = {
    kind: "Workbook",
    props: { name: "book" },
    children,
    parent: null,
    container: null
  };
  children.forEach((child) => {
    child.parent = descriptor;
  });
  return descriptor;
}

describe("renderer commit log helpers", () => {
  it("collects mount ops for workbooks, sheets, and cells", () => {
    const firstCell = cell("A1", 10);
    const secondCell = cell("B1", undefined, "A1*2");
    const root = workbook([sheet("Sheet1", [firstCell, secondCell])]);

    expect(collectMountOps(root)).toEqual([
      { kind: "upsertWorkbook", name: "book" },
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 10 },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "B1", formula: "A1*2" }
    ]);
    expect(collectMountOps(root.children[0]!)).toEqual([
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 10 },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "B1", formula: "A1*2" }
    ]);
    expect(collectMountOps(firstCell)).toEqual([{ kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 10 }]);
  });

  it("collects delete and ordering ops and normalizes duplicate keys", () => {
    const firstSheet = sheet("Sheet1", [cell("A1", 1)]);
    const secondSheet = sheet("Sheet2", [cell("B1", 2)]);
    const root = workbook([firstSheet, secondSheet]);

    expect(collectDeleteOps(root)).toEqual([
      { kind: "deleteSheet", name: "Sheet2" },
      { kind: "deleteSheet", name: "Sheet1" }
    ]);
    expect(collectDeleteOps(firstSheet)).toEqual([{ kind: "deleteSheet", name: "Sheet1" }]);
    expect(collectDeleteOps(firstSheet.children[0]!)).toEqual([{ kind: "deleteCell", sheetName: "Sheet1", addr: "A1" }]);
    expect(collectSheetOrderOps(root)).toEqual([
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "upsertSheet", name: "Sheet2", order: 1 }
    ]);

    expect(
      normalizeCommitOps([
        { kind: "upsertWorkbook", name: "before" },
        { kind: "upsertWorkbook", name: "after" },
        { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 1 },
        { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 2 },
        { kind: "deleteSheet", name: "Sheet2" }
      ])
    ).toEqual([
      { kind: "upsertWorkbook", name: "after" },
      { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 2 },
      { kind: "deleteSheet", name: "Sheet2" }
    ]);
  });
});
