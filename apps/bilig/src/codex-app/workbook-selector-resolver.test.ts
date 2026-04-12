import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import {
  resolveWorkbookSelector,
  resolveWorkbookSelectorToSingleRange,
  type WorkbookSemanticSelector,
  WorkbookSelectorResolutionError,
} from "./workbook-selector-resolver.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";

async function createRuntime(): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: "doc-1",
    replicaId: "server:test",
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", "Revenue");
  engine.setCellValue("Sheet1", "B1", "Margin");
  engine.setCellValue("Sheet1", "A2", 10);
  engine.setCellValue("Sheet1", "B2", 2);
  engine.setCellValue("Sheet1", "A3", 12);
  engine.setCellValue("Sheet1", "B3", 3);
  engine.setCellValue("Sheet1", "D5", "island");
  engine.setDefinedName("Inputs", {
    kind: "range-ref",
    sheetName: "Sheet1",
    startAddress: "A2",
    endAddress: "B3",
  });
  engine.setDefinedName("MarginColumn", {
    kind: "structured-ref",
    tableName: "RevenueTable",
    columnName: "Margin",
  });
  engine.setTable({
    name: "RevenueTable",
    sheetName: "Sheet1",
    startAddress: "A1",
    endAddress: "B3",
    columnNames: ["Revenue", "Margin"],
    headerRow: true,
    totalsRow: false,
  });
  return {
    documentId: "doc-1",
    engine,
    projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
      revision: 7,
      calculatedRevision: 7,
      ownerUserId: "alex@example.com",
      updatedBy: "alex@example.com",
      updatedAt: "2026-04-12T12:00:00.000Z",
    }),
    headRevision: 7,
    calculatedRevision: 7,
    ownerUserId: "alex@example.com",
  };
}

describe("workbook selector resolver", () => {
  it("resolves named ranges, tables, and structured-reference named ranges", async () => {
    const runtime = await createRuntime();

    const namedRange = resolveWorkbookSelectorToSingleRange({
      runtime,
      selector: {
        kind: "namedRange",
        name: "Inputs",
      },
      uiContext: null,
    });
    expect(namedRange.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "A2",
      endAddress: "B3",
    });
    expect(namedRange.resolution.displayLabel).toBe("Inputs");

    const tableRange = resolveWorkbookSelectorToSingleRange({
      runtime,
      selector: {
        kind: "table",
        table: "RevenueTable",
      },
      uiContext: null,
    });
    expect(tableRange.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
    });
    expect(tableRange.resolution.objectType).toBe("table");

    const structuredRef = resolveWorkbookSelectorToSingleRange({
      runtime,
      selector: {
        kind: "namedRange",
        name: "MarginColumn",
      },
      uiContext: null,
    });
    expect(structuredRef.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "B2",
      endAddress: "B3",
    });
    expect(structuredRef.resolution.objectType).toBe("tableColumn");
  });

  it("resolves currentRegion and visibleRows from browser context", async () => {
    const runtime = await createRuntime();

    const currentRegion = resolveWorkbookSelectorToSingleRange({
      runtime,
      selector: {
        kind: "currentRegion",
        anchor: {
          sheet: "Sheet1",
          address: "A2",
        },
      },
      uiContext: null,
    });
    expect(currentRegion.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
    });

    const visibleRows = resolveWorkbookSelectorToSingleRange({
      runtime,
      selector: {
        kind: "visibleRows",
        sheet: "Sheet1",
      },
      uiContext: {
        selection: {
          sheetName: "Sheet1",
          address: "B2",
        },
        viewport: {
          rowStart: 1,
          rowEnd: 2,
          colStart: 0,
          colEnd: 1,
        },
      },
    });
    expect(visibleRows.range).toEqual({
      sheetName: "Sheet1",
      startAddress: "A2",
      endAddress: "D3",
    });
  });

  it("rejects stale selector revisions and scalar named ranges", async () => {
    const runtime = await createRuntime();
    runtime.engine.setDefinedName("ScalarOnly", { kind: "scalar", value: 42 });

    expect(() =>
      resolveWorkbookSelector({
        runtime,
        selector: {
          kind: "namedRange",
          name: "Inputs",
          revision: 6,
        } satisfies WorkbookSemanticSelector,
        uiContext: null,
      }),
    ).toThrowError(WorkbookSelectorResolutionError);

    expect(() =>
      resolveWorkbookSelectorToSingleRange({
        runtime,
        selector: {
          kind: "namedRange",
          name: "ScalarOnly",
        },
        uiContext: null,
      }),
    ).toThrowError(WorkbookSelectorResolutionError);
  });
});
