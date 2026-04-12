import { describe, expect, it } from "vitest";
import {
  appendWorkbookAgentCommandToBundle,
  buildWorkbookAgentExecutionRecord,
  projectWorkbookAgentBundle,
  type WorkbookAgentContextRef,
} from "../workbook-agent-bundles.js";

const selectionContext: WorkbookAgentContextRef = {
  selection: {
    sheetName: "Sheet1",
    address: "B2",
  },
  viewport: {
    rowStart: 0,
    rowEnd: 20,
    colStart: 0,
    colEnd: 10,
  },
};

const rangeSelectionContext: WorkbookAgentContextRef = {
  selection: {
    sheetName: "Sheet1",
    address: "B2",
    range: {
      startAddress: "B2",
      endAddress: "D4",
    },
  },
  viewport: {
    rowStart: 0,
    rowEnd: 20,
    colStart: 0,
    colEnd: 10,
  },
};

describe("workbook agent bundle semantics", () => {
  it("marks selection-only formatting bundles as low-risk selection work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Format the selected cell",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "formatRange",
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B2",
        },
        patch: {
          font: {
            bold: true,
          },
        },
      },
      now: 100,
    });

    expect(bundle.riskClass).toBe("low");
    expect(bundle.scope).toBe("selection");
  });

  it("treats a multi-cell selected range as selection-scoped when the command matches it", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Format the selected range",
      baseRevision: 3,
      context: rangeSelectionContext,
      command: {
        kind: "formatRange",
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "D4",
        },
        patch: {
          fill: "#dbeafe",
        },
      },
      now: 100,
    });

    expect(bundle.scope).toBe("selection");
  });

  it("marks workbook-structure bundles as workbook-scoped structural work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Create a summary sheet",
      baseRevision: 3,
      context: null,
      command: {
        kind: "createSheet",
        name: "Summary",
      },
      now: 100,
    });

    expect(bundle.riskClass).toBe("high");
    expect(bundle.scope).toBe("workbook");
  });

  it("marks sheet deletion as workbook-scoped structural work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Delete the imports sheet",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "deleteSheet",
        name: "Imports",
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Delete sheet Imports",
        riskClass: "high",
        scope: "workbook",
        estimatedAffectedCells: null,
        affectedRanges: [],
      }),
    );
  });

  it("marks non-structural content edits as sheet-scoped work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Write a formula",
      baseRevision: 3,
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "A1",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
      command: {
        kind: "writeRange",
        sheetName: "Sheet1",
        startAddress: "C3",
        values: [[{ formula: "=SUM(A1:B2)" }]],
      },
      now: 100,
    });

    expect(bundle.riskClass).toBe("medium");
    expect(bundle.scope).toBe("sheet");
  });

  it("marks row metadata edits as sheet-scoped changes", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Hide the subtotal rows",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        startRow: 1,
        count: 2,
        hidden: true,
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Hide rows 2-3 in Sheet1",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: null,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A2",
            endAddress: "A3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("marks row insertion as workbook-scoped structural work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Insert summary rows",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "insertRows",
        sheetName: "Sheet1",
        start: 1,
        count: 2,
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Insert rows 2-3 in Sheet1",
        riskClass: "high",
        scope: "workbook",
        estimatedAffectedCells: null,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A2",
            endAddress: "A3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("marks chart creation as sheet-scoped medium-risk workbook object work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Add a revenue chart",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "upsertChart",
        chart: {
          id: "RevenueChart",
          sheetName: "Dashboard",
          address: "B2",
          source: {
            sheetName: "Data",
            startAddress: "A1",
            endAddress: "B4",
          },
          chartType: "column",
          rows: 12,
          cols: 8,
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set chart RevenueChart at Dashboard!B2",
        riskClass: "medium",
        scope: "workbook",
      }),
    );
    expect(bundle.affectedRanges).toEqual([
      {
        sheetName: "Data",
        startAddress: "A1",
        endAddress: "B4",
        role: "source",
      },
      {
        sheetName: "Dashboard",
        startAddress: "B2",
        endAddress: "I13",
        role: "target",
      },
    ]);
  });

  it("marks image placement as sheet-scoped medium-risk media work", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Add a revenue image",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "upsertImage",
        image: {
          id: "RevenueImage",
          sheetName: "Dashboard",
          address: "C3",
          sourceUrl: "https://example.com/revenue.png",
          rows: 8,
          cols: 5,
          altText: "Revenue image",
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set image RevenueImage at Dashboard!C3",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 40,
      }),
    );
    expect(bundle.affectedRanges).toEqual([
      {
        sheetName: "Dashboard",
        startAddress: "C3",
        endAddress: "G10",
        role: "target",
      },
    ]);
  });

  it("normalizes sort ranges and counts affected cells", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Sort the revenue range",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "setSort",
        range: {
          sheetName: "Sheet1",
          startAddress: "B3",
          endAddress: "A1",
        },
        keys: [{ keyAddress: "B1", direction: "desc" }],
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Sort Sheet1!A1:B3 by B1 desc",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 6,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("summarizes and scopes data validation commands against the normalized target range", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Require a status selection",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "setDataValidation",
        validation: {
          range: {
            sheetName: "Sheet1",
            startAddress: "B3",
            endAddress: "A1",
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
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set data validation on Sheet1!A1:B3",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 6,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("summarizes comment thread commands against a single target cell", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Leave a comment on B3",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "upsertCommentThread",
        thread: {
          threadId: "thread-1",
          sheetName: "Sheet1",
          address: "B3",
          comments: [{ id: "comment-1", body: "Check this total." }],
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set comment thread on Sheet1!B3",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 1,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "B3",
            endAddress: "B3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("summarizes conditional format commands against the normalized target range", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Highlight values over ten",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "upsertConditionalFormat",
        format: {
          id: "cf-1",
          range: {
            sheetName: "Sheet1",
            startAddress: "B3",
            endAddress: "A1",
          },
          rule: {
            kind: "cellIs",
            operator: "greaterThan",
            values: [10],
          },
          style: {
            fill: { backgroundColor: "#ff0000" },
          },
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set conditional format on Sheet1!A1:B3",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 6,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "B3",
            role: "target",
          },
        ],
      }),
    );
  });

  it("treats sheet protection commands as high-risk workbook-scope changes", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Protect Sheet1",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "setSheetProtection",
        protection: {
          sheetName: "Sheet1",
          hideFormulas: true,
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Protect sheet Sheet1 and hide formulas",
        riskClass: "high",
        scope: "workbook",
      }),
    );
  });

  it("summarizes note commands against the normalized target cell", () => {
    const bundle = appendWorkbookAgentCommandToBundle({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Attach a note to C4",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "upsertNote",
        note: {
          sheetName: "Sheet1",
          address: "c4",
          text: "Manual override",
        },
      },
      now: 100,
    });

    expect(bundle).toEqual(
      expect.objectContaining({
        summary: "Set note on Sheet1!C4",
        riskClass: "medium",
        scope: "sheet",
        estimatedAffectedCells: 1,
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "C4",
            endAddress: "C4",
            role: "target",
          },
        ],
      }),
    );
  });

  it("projects a scoped subset as its own preview/apply bundle", () => {
    const staged = appendWorkbookAgentCommandToBundle({
      previousBundle: appendWorkbookAgentCommandToBundle({
        previousBundle: null,
        documentId: "doc-1",
        threadId: "thr-1",
        turnId: "turn-1",
        goalText: "Update two cells",
        baseRevision: 3,
        context: selectionContext,
        command: {
          kind: "writeRange",
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[1]],
        },
        now: 100,
      }),
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Update two cells",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "writeRange",
        sheetName: "Sheet1",
        startAddress: "C3",
        values: [[2]],
      },
      now: 100,
    });

    const subset = projectWorkbookAgentBundle({
      bundle: staged,
      commandIndexes: [1],
      bundleId: staged.id,
    });

    expect(subset).toEqual(
      expect.objectContaining({
        id: staged.id,
        summary: "Write cells in Sheet1!C3",
        commands: [
          {
            kind: "writeRange",
            sheetName: "Sheet1",
            startAddress: "C3",
            values: [[2]],
          },
        ],
        estimatedAffectedCells: 1,
      }),
    );
  });

  it("records partial acceptance with only the applied command subset", () => {
    const staged = appendWorkbookAgentCommandToBundle({
      previousBundle: appendWorkbookAgentCommandToBundle({
        previousBundle: null,
        documentId: "doc-1",
        threadId: "thr-1",
        turnId: "turn-1",
        goalText: "Update two cells",
        baseRevision: 3,
        context: selectionContext,
        command: {
          kind: "writeRange",
          sheetName: "Sheet1",
          startAddress: "B2",
          values: [[1]],
        },
        now: 100,
      }),
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Update two cells",
      baseRevision: 3,
      context: selectionContext,
      command: {
        kind: "writeRange",
        sheetName: "Sheet1",
        startAddress: "C3",
        values: [[2]],
      },
      now: 100,
    });
    const subset = projectWorkbookAgentBundle({
      bundle: staged,
      commandIndexes: [1],
      bundleId: staged.id,
    });
    if (!subset) {
      throw new Error("Expected a projected subset bundle");
    }

    const record = buildWorkbookAgentExecutionRecord({
      bundle: subset,
      actorUserId: "alex@example.com",
      planText: "Update the target cell only",
      preview: null,
      appliedRevision: 4,
      appliedBy: "user",
      acceptedScope: "partial",
      now: 200,
    });

    expect(record.bundleId).toBe(staged.id);
    expect(record.acceptedScope).toBe("partial");
    expect(record.summary).toBe("Write cells in Sheet1!C3");
    expect(record.commands).toEqual([
      {
        kind: "writeRange",
        sheetName: "Sheet1",
        startAddress: "C3",
        values: [[2]],
      },
    ]);
  });
});
