import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import { buildWorkbookAgentPreview } from "../workbook-agent-preview.js";
import {
  decodeWorkbookAgentPreviewSummary,
  type WorkbookAgentCommandBundle,
} from "../workbook-agent-bundles.js";

async function createSnapshot() {
  const engine = new SpreadsheetEngine({
    workbookName: "Preview Workbook",
    replicaId: "test:preview",
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 42);
  engine.setCellFormula("Sheet1", "B1", "SUM(A1:A1)");
  return engine.exportSnapshot();
}

function createBundle(
  command: WorkbookAgentCommandBundle["commands"][number],
): WorkbookAgentCommandBundle {
  return {
    id: "bundle-1",
    documentId: "doc-1",
    threadId: "thr-1",
    turnId: "turn-1",
    goalText: "Preview workbook changes",
    summary: "Preview workbook changes",
    scope: "sheet",
    riskClass: "medium",
    approvalMode: "preview",
    baseRevision: 1,
    createdAtUnixMs: 1,
    context: null,
    commands: [command],
    affectedRanges:
      command.kind === "formatRange"
        ? [
            {
              sheetName: command.range.sheetName,
              startAddress: command.range.startAddress,
              endAddress: command.range.endAddress,
              role: "target",
            },
          ]
        : [
            {
              sheetName: "Sheet1",
              startAddress: "B1",
              endAddress: "B1",
              role: "target",
            },
          ],
    estimatedAffectedCells: 1,
  };
}

describe("workbook agent preview", () => {
  it("captures formula and value change kinds in sampled diffs", async () => {
    const preview = await buildWorkbookAgentPreview({
      snapshot: await createSnapshot(),
      replicaId: "preview",
      bundle: createBundle({
        kind: "writeRange",
        sheetName: "Sheet1",
        startAddress: "B1",
        values: [[{ formula: "=A1*2" }]],
      }),
    });

    expect(preview.cellDiffs).toEqual([
      expect.objectContaining({
        sheetName: "Sheet1",
        address: "B1",
        beforeFormula: "=SUM(A1:A1)",
        afterFormula: "=A1*2",
        changeKinds: ["formula"],
      }),
    ]);
    expect(preview.effectSummary.formulaChangeCount).toBe(1);
    expect(preview.effectSummary.inputChangeCount).toBe(0);
  });

  it("captures style and number-format-only preview diffs", async () => {
    const preview = await buildWorkbookAgentPreview({
      snapshot: await createSnapshot(),
      replicaId: "preview",
      bundle: createBundle({
        kind: "formatRange",
        range: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
        },
        patch: {
          font: {
            bold: true,
          },
        },
        numberFormat: "currency",
      }),
    });

    expect(preview.cellDiffs).toEqual([
      expect.objectContaining({
        sheetName: "Sheet1",
        address: "A1",
        changeKinds: expect.arrayContaining(["style", "numberFormat"]),
      }),
    ]);
    expect(preview.effectSummary.styleChangeCount).toBe(1);
    expect(preview.effectSummary.numberFormatChangeCount).toBe(1);
  });

  it("captures structural row metadata previews without cell diffs", async () => {
    const preview = await buildWorkbookAgentPreview({
      snapshot: await createSnapshot(),
      replicaId: "preview",
      bundle: {
        id: "bundle-rows",
        documentId: "doc-1",
        threadId: "thr-1",
        turnId: "turn-1",
        goalText: "Hide subtotal rows",
        summary: "Hide subtotal rows",
        scope: "sheet",
        riskClass: "medium",
        approvalMode: "preview",
        baseRevision: 1,
        createdAtUnixMs: 1,
        context: null,
        commands: [
          {
            kind: "updateRowMetadata",
            sheetName: "Sheet1",
            startRow: 1,
            count: 2,
            hidden: true,
          },
        ],
        affectedRanges: [
          {
            sheetName: "Sheet1",
            startAddress: "A2",
            endAddress: "A3",
            role: "target",
          },
        ],
        estimatedAffectedCells: null,
      },
    });

    expect(preview.structuralChanges).toEqual(["Hide rows 2-3 in Sheet1"]);
    expect(preview.cellDiffs).toEqual([]);
    expect(preview.effectSummary.structuralChangeCount).toBe(1);
    expect(preview.effectSummary.displayedCellDiffCount).toBe(0);
  });

  it("normalizes legacy preview payloads without dropping persisted runs", () => {
    const decoded = decodeWorkbookAgentPreviewSummary({
      ranges: [],
      structuralChanges: [],
      cellDiffs: [
        {
          sheetName: "Sheet1",
          address: "A1",
          beforeInput: 1,
          beforeFormula: null,
          afterInput: 2,
          afterFormula: null,
        },
      ],
    });

    expect(decoded).toEqual(
      expect.objectContaining({
        cellDiffs: [
          expect.objectContaining({
            changeKinds: ["input"],
          }),
        ],
        effectSummary: expect.objectContaining({
          displayedCellDiffCount: 1,
          inputChangeCount: 1,
          structuralChangeCount: 0,
        }),
      }),
    );
  });
});
