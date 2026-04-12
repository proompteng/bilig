import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import {
  createWorkbookMetadataService,
  type WorkbookMetadataService,
} from "../workbook-metadata-service.js";
import { createWorkbookMetadataRecord } from "../workbook-metadata-types.js";

function createService(): WorkbookMetadataService {
  return createWorkbookMetadataService(createWorkbookMetadataRecord());
}

describe("WorkbookMetadataService", () => {
  it("clones caller-owned defined name objects on write and read", () => {
    const service = createService();
    const source = {
      kind: "range-ref" as const,
      sheetName: "Data",
      startAddress: "A1",
      endAddress: "B4",
    };

    const stored = Effect.runSync(service.setDefinedName(" SalesRange ", source));
    source.sheetName = "Mutated";
    if (stored.value && typeof stored.value === "object" && stored.value.kind === "range-ref") {
      stored.value.startAddress = "Z9";
    }

    expect(Effect.runSync(service.getDefinedName("salesrange"))).toEqual({
      name: "SalesRange",
      value: {
        kind: "range-ref",
        sheetName: "Data",
        startAddress: "A1",
        endAddress: "B4",
      },
    });
    expect(Effect.runSync(service.listDefinedNames())).toEqual([
      {
        name: "SalesRange",
        value: {
          kind: "range-ref",
          sheetName: "Data",
          startAddress: "A1",
          endAddress: "B4",
        },
      },
    ]);
  });

  it("clones and normalizes data validation records on write and read", () => {
    const service = createService();
    const input = {
      range: {
        sheetName: "Sheet1",
        startAddress: "c4",
        endAddress: "b2",
      },
      rule: {
        kind: "list" as const,
        values: ["Draft", "Final"],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: "stop" as const,
      errorTitle: "Status required",
      errorMessage: "Pick Draft or Final.",
    };

    const stored = Effect.runSync(service.setDataValidation(input));
    input.rule.values[0] = "Mutated";
    stored.rule.values.push("Broken");

    expect(
      Effect.runSync(
        service.getDataValidation("Sheet1", {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "C4",
        }),
      ),
    ).toEqual({
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "C4",
      },
      rule: {
        kind: "list",
        values: ["Draft", "Final"],
      },
      allowBlank: false,
      showDropdown: true,
      errorStyle: "stop",
      errorTitle: "Status required",
      errorMessage: "Pick Draft or Final.",
    });
  });

  it("clones and normalizes comment threads and notes on write and read", () => {
    const service = createService();
    const threadInput = {
      threadId: " thread-1 ",
      sheetName: "Sheet1",
      address: "c4",
      comments: [{ id: " comment-1 ", body: "Check this total." }],
    };
    const noteInput = {
      sheetName: "Sheet1",
      address: "d5",
      text: " Manual override ",
    };

    const storedThread = Effect.runSync(service.setCommentThread(threadInput));
    const storedNote = Effect.runSync(service.setNote(noteInput));
    threadInput.comments[0].body = "Mutated";
    storedThread.comments[0].body = "Leaked";
    noteInput.text = "Changed";
    storedNote.text = "Broken";

    expect(Effect.runSync(service.getCommentThread("Sheet1", "C4"))).toEqual({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "C4",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    expect(Effect.runSync(service.getNote("Sheet1", "D5"))).toEqual({
      sheetName: "Sheet1",
      address: "D5",
      text: "Manual override",
    });
  });

  it("normalizes and clones pivot records so caller mutation does not leak back into metadata", () => {
    const service = createService();
    const input = {
      name: " RevenuePivot ",
      sheetName: "Sheet1",
      address: "c3",
      source: { sheetName: "Data", startAddress: "b4", endAddress: "a1" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" as const }],
      rows: 3,
      cols: 2,
    };

    const stored = Effect.runSync(service.setPivot(input));
    input.groupBy[0] = "Mutated";
    input.values[0].summarizeBy = "count";
    stored.groupBy.push("Leaked");
    stored.values[0].field = "Broken";
    stored.source.startAddress = "Z9";

    expect(Effect.runSync(service.getPivot("Sheet1", "C3"))).toEqual({
      name: "RevenuePivot",
      sheetName: "Sheet1",
      address: "C3",
      source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });
    expect(Effect.runSync(service.listPivots())).toEqual([
      {
        name: "RevenuePivot",
        sheetName: "Sheet1",
        address: "C3",
        source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
        groupBy: ["Region"],
        values: [{ field: "Sales", summarizeBy: "sum" }],
        rows: 3,
        cols: 2,
      },
    ]);
  });

  it("renames, deletes, and resets workbook metadata across buckets", () => {
    const metadata = createWorkbookMetadataRecord();
    const service = createWorkbookMetadataService(metadata);

    Effect.runSync(service.setFreezePane("Source", 1, 2));
    Effect.runSync(
      service.setFilter("Source", {
        sheetName: "Source",
        startAddress: "C3",
        endAddress: "A1",
      }),
    );
    Effect.runSync(
      service.setSort("Source", { sheetName: "Source", startAddress: "C3", endAddress: "A1" }, [
        { keyAddress: "B1", direction: "asc" },
      ]),
    );
    Effect.runSync(
      service.setDataValidation({
        range: { sheetName: "Source", startAddress: "B4", endAddress: "A2" },
        rule: {
          kind: "list",
          values: ["Draft", "Final"],
        },
        allowBlank: false,
        showDropdown: true,
      }),
    );
    Effect.runSync(
      service.setCommentThread({
        threadId: "thread-1",
        sheetName: "Source",
        address: "C4",
        comments: [{ id: "comment-1", body: "Check this total." }],
      }),
    );
    Effect.runSync(
      service.setNote({
        sheetName: "Source",
        address: "D5",
        text: "Manual override",
      }),
    );
    Effect.runSync(
      service.setTable({
        name: " Revenue ",
        sheetName: "Source",
        startAddress: "A1",
        endAddress: "C10",
        columnNames: ["Region", "Sales"],
        headerRow: true,
        totalsRow: false,
      }),
    );
    Effect.runSync(service.setSpill("Source", "b2", 2, 3));
    Effect.runSync(
      service.setPivot({
        name: " RevenuePivot ",
        sheetName: "Source",
        address: "c3",
        source: { sheetName: "Source", startAddress: "b4", endAddress: "a1" },
        groupBy: ["Region"],
        values: [{ field: "Sales", summarizeBy: "sum" }],
        rows: 3,
        cols: 2,
      }),
    );
    Effect.runSync(service.setWorkbookProperty(" Author ", "greg"));
    Effect.runSync(
      service.setDefinedName(" LocalName ", { kind: "formula", formula: "=Source!A1" }),
    );
    Effect.runSync(service.setCalculationSettings({ mode: "manual" }));
    Effect.runSync(service.setVolatileContext({ recalcEpoch: 7 }));
    metadata.rowMetadata.set("Source:0:2", {
      sheetName: "Source",
      start: 0,
      count: 2,
      size: 24,
      hidden: null,
    });
    metadata.columnMetadata.set("Source:1:1", {
      sheetName: "Source",
      start: 1,
      count: 1,
      size: null,
      hidden: true,
    });

    Effect.runSync(service.renameSheet("Source", "Renamed"));

    expect(Effect.runSync(service.getFreezePane("Renamed"))).toEqual({
      sheetName: "Renamed",
      rows: 1,
      cols: 2,
    });
    expect(
      Effect.runSync(
        service.getFilter("Renamed", {
          sheetName: "Renamed",
          startAddress: "A1",
          endAddress: "C3",
        }),
      ),
    ).toEqual({
      sheetName: "Renamed",
      range: { sheetName: "Renamed", startAddress: "A1", endAddress: "C3" },
    });
    expect(
      Effect.runSync(
        service.getSort("Renamed", {
          sheetName: "Renamed",
          startAddress: "A1",
          endAddress: "C3",
        }),
      ),
    ).toEqual({
      sheetName: "Renamed",
      range: { sheetName: "Renamed", startAddress: "A1", endAddress: "C3" },
      keys: [{ keyAddress: "B1", direction: "asc" }],
    });
    expect(
      Effect.runSync(
        service.getDataValidation("Renamed", {
          sheetName: "Renamed",
          startAddress: "A2",
          endAddress: "B4",
        }),
      ),
    ).toEqual({
      range: { sheetName: "Renamed", startAddress: "A2", endAddress: "B4" },
      rule: {
        kind: "list",
        values: ["Draft", "Final"],
      },
      allowBlank: false,
      showDropdown: true,
    });
    expect(Effect.runSync(service.getCommentThread("Renamed", "C4"))).toEqual({
      threadId: "thread-1",
      sheetName: "Renamed",
      address: "C4",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    expect(Effect.runSync(service.getNote("Renamed", "D5"))).toEqual({
      sheetName: "Renamed",
      address: "D5",
      text: "Manual override",
    });
    expect(Effect.runSync(service.getSpill("Renamed", "B2"))).toEqual({
      sheetName: "Renamed",
      address: "B2",
      rows: 2,
      cols: 3,
    });
    expect(Effect.runSync(service.getPivot("Renamed", "C3"))).toEqual({
      name: "RevenuePivot",
      sheetName: "Renamed",
      address: "C3",
      source: { sheetName: "Renamed", startAddress: "A1", endAddress: "B4" },
      groupBy: ["Region"],
      values: [{ field: "Sales", summarizeBy: "sum" }],
      rows: 3,
      cols: 2,
    });
    expect(Effect.runSync(service.getTable("Revenue"))).toEqual({
      name: "Revenue",
      sheetName: "Renamed",
      startAddress: "A1",
      endAddress: "C10",
      columnNames: ["Region", "Sales"],
      headerRow: true,
      totalsRow: false,
    });
    expect([...metadata.rowMetadata.values()]).toEqual([
      { sheetName: "Renamed", start: 0, count: 2, size: 24, hidden: null },
    ]);
    expect([...metadata.columnMetadata.values()]).toEqual([
      { sheetName: "Renamed", start: 1, count: 1, size: null, hidden: true },
    ]);

    Effect.runSync(service.deleteSheetRecords("Renamed"));

    expect(Effect.runSync(service.getFreezePane("Renamed"))).toBeUndefined();
    expect(Effect.runSync(service.listFilters("Renamed"))).toEqual([]);
    expect(Effect.runSync(service.listSorts("Renamed"))).toEqual([]);
    expect(Effect.runSync(service.listDataValidations("Renamed"))).toEqual([]);
    expect(Effect.runSync(service.listCommentThreads("Renamed"))).toEqual([]);
    expect(Effect.runSync(service.listNotes("Renamed"))).toEqual([]);
    expect(Effect.runSync(service.listPivots())).toEqual([]);
    expect(Effect.runSync(service.listTables())).toEqual([]);
    expect(Effect.runSync(service.listSpills())).toEqual([]);
    expect([...metadata.rowMetadata.values()]).toEqual([]);
    expect([...metadata.columnMetadata.values()]).toEqual([]);
    expect(Effect.runSync(service.getWorkbookProperty("Author"))).toEqual({
      key: "Author",
      value: "greg",
    });
    expect(Effect.runSync(service.getDefinedName("LocalName"))).toEqual({
      name: "LocalName",
      value: { kind: "formula", formula: "=Source!A1" },
    });

    Effect.runSync(service.reset());

    expect(Effect.runSync(service.listWorkbookProperties())).toEqual([]);
    expect(Effect.runSync(service.listDefinedNames())).toEqual([]);
    expect(Effect.runSync(service.getCalculationSettings())).toEqual({
      mode: "automatic",
      compatibilityMode: "excel-modern",
    });
    expect(Effect.runSync(service.getVolatileContext())).toEqual({
      recalcEpoch: 0,
    });
  });
});
