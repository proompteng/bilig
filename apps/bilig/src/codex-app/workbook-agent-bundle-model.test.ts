import { describe, expect, it } from "vitest";
import { appendWorkbookAgentBundleCommand } from "./workbook-agent-bundle-model.js";

describe("workbook agent bundle model", () => {
  it("marks selection-only formatting bundles as auto-apply", () => {
    const bundle = appendWorkbookAgentBundleCommand({
      previousBundle: null,
      documentId: "doc-1",
      threadId: "thr-1",
      turnId: "turn-1",
      goalText: "Format the selected cell",
      baseRevision: 3,
      context: {
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
      },
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

  it("marks workbook-structure bundles as explicit approval", () => {
    const bundle = appendWorkbookAgentBundleCommand({
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
    const bundle = appendWorkbookAgentBundleCommand({
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
});
