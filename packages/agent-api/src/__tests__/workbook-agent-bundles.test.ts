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
  it("marks selection-only formatting bundles as auto-apply", () => {
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
    expect(bundle.approvalMode).toBe("auto");
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
    expect(bundle.approvalMode).toBe("auto");
  });

  it("marks workbook-structure bundles as explicit approval", () => {
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
    expect(bundle.approvalMode).toBe("explicit");
  });

  it("marks non-structural content edits as preview-required", () => {
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
    expect(bundle.approvalMode).toBe("preview");
  });

  it("marks row metadata edits as sheet-scoped preview bundles", () => {
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
        approvalMode: "preview",
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
        approvalMode: "explicit",
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
