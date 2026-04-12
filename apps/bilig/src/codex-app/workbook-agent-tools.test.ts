import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import { handleWorkbookAgentToolCall } from "./workbook-agent-tools.js";
import { z } from "zod";

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: "doc-1",
    replicaId: "server:test",
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 42);
  engine.setCellValue("Sheet1", "A2", "Gross Margin");
  engine.setCellFormula("Sheet1", "B1", "SUM(A1:A1)");
  engine.setCellFormula("Sheet1", "C1", "1/0");
  engine.setCellFormula("Sheet1", "D1", "LEN(A1:A2)");
  engine.setRangeStyle(
    { sheetName: "Sheet1", startAddress: "A1", endAddress: "B1" },
    {
      fill: { backgroundColor: "#fef3c7" },
      font: { bold: true },
    },
  );
  engine.setRangeNumberFormat(
    { sheetName: "Sheet1", startAddress: "A1", endAddress: "B1" },
    {
      kind: "currency",
      currency: "USD",
      decimals: 2,
      useGrouping: true,
      negativeStyle: "minus",
      zeroStyle: "zero",
    },
  );
  engine.createSheet("Ops Search");
  engine.setCellValue("Ops Search", "A1", "Northwind Import");
  return engine;
}

function createZeroSyncHarness(engine: SpreadsheetEngine) {
  const zeroSyncService: ZeroSyncService = {
    enabled: true,
    async initialize() {},
    async close() {},
    async handleQuery() {
      throw new Error("not used");
    },
    async handleMutate() {
      throw new Error("not used");
    },
    async inspectWorkbook(_documentId, task) {
      const runtime: WorkbookRuntime = {
        documentId: "doc-1",
        engine,
        projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
          revision: 1,
          calculatedRevision: 1,
          ownerUserId: "alex@example.com",
          updatedBy: "alex@example.com",
          updatedAt: "2026-04-06T12:00:00.000Z",
        }),
        headRevision: 1,
        calculatedRevision: 1,
        ownerUserId: "alex@example.com",
      };
      return await task(runtime);
    },
    async applyServerMutator() {
      throw new Error("not used");
    },
    async applyAgentCommandBundle() {
      return {
        revision: 1,
        preview: {
          ranges: [],
          structuralChanges: [],
          cellDiffs: [],
          effectSummary: {
            displayedCellDiffCount: 0,
            truncatedCellDiffs: false,
            inputChangeCount: 0,
            formulaChangeCount: 0,
            styleChangeCount: 0,
            numberFormatChangeCount: 0,
            structuralChangeCount: 0,
          },
        },
      };
    },
    async listWorkbookChanges() {
      return [];
    },
    async listWorkbookAgentRuns() {
      return [];
    },
    async listWorkbookAgentThreadRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {
      throw new Error("not used");
    },
    async listWorkbookAgentThreadSummaries() {
      return [];
    },
    async loadWorkbookAgentThreadState() {
      return null;
    },
    async saveWorkbookAgentThreadState() {
      throw new Error("not used");
    },
    async listWorkbookThreadWorkflowRuns() {
      return [];
    },
    async upsertWorkbookWorkflowRun() {
      throw new Error("not used");
    },
    async getWorkbookHeadRevision() {
      return 1;
    },
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
  };
  return { zeroSyncService };
}

function createBundle(
  command: WorkbookAgentCommandBundle["commands"][number],
): WorkbookAgentCommandBundle {
  return {
    id: "bundle-1",
    documentId: "doc-1",
    threadId: "thr-1",
    turnId: "turn-1",
    goalText: "Update cells",
    summary: "Stage workbook changes",
    scope: "selection",
    riskClass: "low",
    baseRevision: 1,
    createdAtUnixMs: 1,
    context: null,
    commands: [command],
    affectedRanges: [],
    estimatedAffectedCells: 2,
  };
}

const workbookSummarySchema = z.object({
  summary: z.object({
    sheetCount: z.number(),
    tableCount: z.number(),
    pivotCount: z.number(),
    spillCount: z.number(),
    hiddenRowIndexCount: z.number(),
    hiddenColumnIndexCount: z.number(),
  }),
  sheets: z.array(
    z.object({
      name: z.string(),
      freezePane: z
        .object({
          rows: z.number(),
          cols: z.number(),
        })
        .nullable(),
      filterCount: z.number(),
      sortCount: z.number(),
      tableCount: z.number(),
      pivotCount: z.number(),
      spillCount: z.number(),
      rowMetadata: z.object({
        hiddenIndexCount: z.number(),
        explicitSizeIndexCount: z.number(),
      }),
      columnMetadata: z.object({
        hiddenIndexCount: z.number(),
        explicitSizeIndexCount: z.number(),
      }),
    }),
  ),
});

describe("workbook agent tools", () => {
  it("reads workbook structure with sheet metadata for workbook-wide prompts", async () => {
    const engine = await createEngine();
    engine.updateRowMetadata("Sheet1", 1, 2, 24, true);
    engine.updateColumnMetadata("Sheet1", 0, 1, 110, true);
    engine.setFreezePane("Sheet1", 1, 0);
    engine.setFilter("Sheet1", { sheetName: "Sheet1", startAddress: "A1", endAddress: "D3" });
    engine.setSort("Sheet1", { sheetName: "Sheet1", startAddress: "A1", endAddress: "D3" }, [
      { keyAddress: "B1", direction: "desc" },
    ]);
    engine.setTable({
      name: "Sheet1Table",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "D3",
      columnNames: ["Revenue", "Formula", "Error", "Length"],
      headerRow: true,
      totalsRow: false,
    });
    engine.setSpillRange("Sheet1", "F1", 2, 2);
    engine.setPivotTable("Ops Search", "B2", {
      name: "ImportPivot",
      source: { sheetName: "Sheet1", startAddress: "A1", endAddress: "D3" },
      groupBy: ["Revenue"],
      values: [{ sourceColumn: "Formula", summarizeBy: "count" }],
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-workbook",
        tool: "bilig_read_workbook",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const payload = workbookSummarySchema.parse(
      JSON.parse(textItem && "text" in textItem ? textItem.text : ""),
    );
    expect(payload.summary).toEqual(
      expect.objectContaining({
        sheetCount: 2,
        tableCount: 1,
        pivotCount: 1,
        spillCount: 1,
        hiddenRowIndexCount: 2,
        hiddenColumnIndexCount: 1,
      }),
    );
    expect(payload.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: "Sheet1",
          freezePane: { rows: 1, cols: 0 },
          filterCount: 1,
          sortCount: 1,
          tableCount: 1,
          spillCount: 1,
          rowMetadata: expect.objectContaining({
            hiddenIndexCount: 2,
            explicitSizeIndexCount: 2,
          }),
          columnMetadata: expect.objectContaining({
            hiddenIndexCount: 1,
            explicitSizeIndexCount: 1,
          }),
        }),
        expect.objectContaining({
          name: "Ops Search",
          pivotCount: 1,
        }),
      ]),
    );
  });

  it("reads the current browser selection through the attached workbook context", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
            range: {
              startAddress: "B2",
              endAddress: "D5",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-selection",
        tool: "bilig_read_selection",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"startAddress": "B2"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"endAddress": "D5"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"styleId"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"sheetState"');
  });

  it("reads workbook context with selection geometry and sheet state", async () => {
    const engine = await createEngine();
    engine.updateRowMetadata("Sheet1", 1, 2, 28, true);
    engine.updateColumnMetadata("Sheet1", 2, 1, 140, true);
    engine.setFreezePane("Sheet1", 1, 0);
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
            range: {
              startAddress: "B2",
              endAddress: "D5",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-context",
        tool: "bilig_get_context",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"kind": "range"');
    expect(text).toContain('"cellCount": 12');
    expect(text).toContain('"freezePane": {');
    expect(text).toContain('"hiddenRows"');
    expect(text).toContain('"hiddenColumns"');
  });

  it("lists sheets and reads sheet-level workbook view metadata", async () => {
    const engine = await createEngine();
    engine.setFreezePane("Sheet1", 1, 0);
    engine.setFilter("Sheet1", { sheetName: "Sheet1", startAddress: "A1", endAddress: "D3" });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const sheetsResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-sheets",
        tool: "list_sheets",
        arguments: {},
      },
    );

    expect(sheetsResponse.success).toBe(true);
    const sheetsText = sheetsResponse.contentItems[0];
    expect(sheetsText?.type).toBe("inputText");
    expect(sheetsText && "text" in sheetsText ? sheetsText.text : "").toContain('"name": "Sheet1"');
    expect(sheetsText && "text" in sheetsText ? sheetsText.text : "").toContain(
      '"name": "Ops Search"',
    );

    const sheetViewResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-sheet-view",
        tool: "get_sheet_view",
        arguments: {
          sheetName: "Sheet1",
        },
      },
    );

    expect(sheetViewResponse.success).toBe(true);
    const sheetViewText = sheetViewResponse.contentItems[0];
    expect(sheetViewText?.type).toBe("inputText");
    expect(sheetViewText && "text" in sheetViewText ? sheetViewText.text : "").toContain(
      '"freezePane": {',
    );
    expect(sheetViewText && "text" in sheetViewText ? sheetViewText.text : "").toContain(
      '"filters": [',
    );
  });

  it("reads used range, current region, and axis metadata", async () => {
    const engine = await createEngine();
    engine.updateRowMetadata("Sheet1", 1, 2, 28, true);
    engine.updateColumnMetadata("Sheet1", 0, 1, 120, true);
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const usedRangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-used-range",
        tool: "get_used_range",
        arguments: {
          sheetName: "Sheet1",
        },
      },
    );

    expect(usedRangeResponse.success).toBe(true);
    const usedRangeText = usedRangeResponse.contentItems[0];
    expect(usedRangeText?.type).toBe("inputText");
    expect(usedRangeText && "text" in usedRangeText ? usedRangeText.text : "").toContain(
      '"startAddress": "A1"',
    );

    const currentRegionResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-current-region",
        tool: "get_current_region",
        arguments: {
          sheetName: "Sheet1",
          address: "A1",
        },
      },
    );

    expect(currentRegionResponse.success).toBe(true);
    const currentRegionText = currentRegionResponse.contentItems[0];
    expect(currentRegionText?.type).toBe("inputText");
    expect(
      currentRegionText && "text" in currentRegionText ? currentRegionText.text : "",
    ).toContain('"derivedA1Ranges": [');

    const rowMetadataResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-row-metadata",
        tool: "get_row_metadata",
        arguments: {
          sheetName: "Sheet1",
        },
      },
    );

    expect(rowMetadataResponse.success).toBe(true);
    const rowMetadataText = rowMetadataResponse.contentItems[0];
    expect(rowMetadataText?.type).toBe("inputText");
    expect(rowMetadataText && "text" in rowMetadataText ? rowMetadataText.text : "").toContain(
      '"hidden": true',
    );

    const columnMetadataResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-column-metadata",
        tool: "get_column_metadata",
        arguments: {
          sheetName: "Sheet1",
        },
      },
    );

    expect(columnMetadataResponse.success).toBe(true);
    const columnMetadataText = columnMetadataResponse.contentItems[0];
    expect(columnMetadataText?.type).toBe("inputText");
    expect(
      columnMetadataText && "text" in columnMetadataText ? columnMetadataText.text : "",
    ).toContain('"size": 120');
  });

  it("reads the visible viewport through the attached browser context", async () => {
    const engine = await createEngine();
    engine.setFreezePane("Sheet1", 1, 0);
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 1,
            colStart: 0,
            colEnd: 1,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-visible",
        tool: "bilig_read_visible_range",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"startAddress": "A1"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"endAddress": "B2"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"freezePane": {');
  });

  it("inspects one cell with formula lineage and runtime metadata", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-inspect",
        tool: "bilig_inspect_cell",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"address": "B1"');
    expect(text).toContain('"formula": "=SUM(A1:A1)"');
    expect(text).toContain('"directPrecedents": [');
    expect(text).toContain("Sheet1!A1");
    expect(text).toContain('"style": {');
    expect(text).toContain('"backgroundColor": "#fef3c7"');
    expect(text).toContain('"numberFormat": {');
  });

  it("scans formula issues through the warm local runtime", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-formula-issues",
        tool: "bilig_find_formula_issues",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"issueCount": 2');
    expect(text).toContain('"address": "C1"');
    expect(text).toContain('"errorText": "#DIV/0!"');
    expect(text).toContain('"address": "D1"');
    expect(text).toContain('"unsupported"');
  });

  it("searches workbook sheets, cells, formulas, and values", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-search",
        tool: "bilig_search_workbook",
        arguments: {
          query: "gross margin",
        },
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"query": "gross margin"');
    expect(text).toContain('"address": "A2"');
    expect(text).toContain('"snippet": "Gross Margin"');
  });

  it("reads recent durable workbook changes", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const listWorkbookChanges = vi.fn(async () => [
      {
        revision: 12,
        actorUserId: "alex@example.com",
        clientMutationId: null,
        eventKind: "applyAgentCommandBundle" as const,
        summary: "Applied workbook change set at revision r12",
        sheetId: 1,
        sheetName: "Sheet1",
        anchorAddress: "B2",
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "C4",
        },
        undoBundle: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 1_234,
      },
    ]);
    zeroSyncService.listWorkbookChanges = listWorkbookChanges;

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-recent-changes",
        tool: "bilig_read_recent_changes",
        arguments: {
          limit: 5,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(listWorkbookChanges).toHaveBeenCalledWith("doc-1", 5);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"changeCount": 1');
    expect(text).toContain('"revision": 12');
    expect(text).toContain('"summary": "Applied workbook change set at revision r12"');
    expect(text).toContain('"startAddress": "B2"');
  });

  it("traces multi-hop workbook dependencies from the attached selection", async () => {
    const engine = await createEngine();
    engine.setCellFormula("Sheet1", "E1", "B1*2");
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-trace",
        tool: "bilig_trace_dependencies",
        arguments: {
          direction: "both",
          depth: 2,
        },
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"address": "B1"');
    expect(text).toContain('"precedentCount": 1');
    expect(text).toContain('"dependentCount": 1');
    expect(text).toContain('"address": "A1"');
    expect(text).toContain('"address": "E1"');
  });

  it("reads workbook ranges through the authoritative runtime", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-1",
        tool: "bilig_read_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "B1",
        },
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"address": "A1"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"value": 42');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain(
      '"formula": "=SUM(A1:A1)"',
    );
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"styles": [');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain(
      '"backgroundColor": "#fef3c7"',
    );
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"numberFormats": [');
  });

  it("reads discontiguous selector results as ordered range sets", async () => {
    const engine = await createEngine();
    engine.setCellValue("Sheet1", "A1", "Revenue");
    engine.setCellValue("Sheet1", "B1", "Margin");
    engine.setCellValue("Sheet1", "A2", 10);
    engine.setCellValue("Sheet1", "B2", 2);
    engine.setCellValue("Sheet1", "A3", 12);
    engine.setCellValue("Sheet1", "B3", 3);
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-row-query",
        tool: "read_range",
        arguments: {
          selector: {
            kind: "rowQuery",
            sheet: "Sheet1",
            predicate: {
              column: "Revenue",
              op: "gte",
              value: 10,
            },
          },
        },
      },
    );

    expect(response.success).toBe(true);
    const textItem = response.contentItems[0];
    expect(textItem?.type).toBe("inputText");
    const text = textItem && "text" in textItem ? textItem.text : "";
    expect(text).toContain('"rangeCount": 2');
    expect(text).toContain('"startAddress": "A2"');
    expect(text).toContain('"endAddress": "D3"');
  });

  it("lists named ranges and tables from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setDefinedName("Inputs", {
      kind: "range-ref",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B1",
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
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const namedRangesResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-named-ranges",
        tool: "list_named_ranges",
        arguments: {},
      },
    );

    expect(namedRangesResponse.success).toBe(true);
    const namedRangesText = namedRangesResponse.contentItems[0];
    expect(namedRangesText?.type).toBe("inputText");
    expect(namedRangesText && "text" in namedRangesText ? namedRangesText.text : "").toContain(
      '"name": "Inputs"',
    );
    expect(namedRangesText && "text" in namedRangesText ? namedRangesText.text : "").toContain(
      '"startAddress": "A1"',
    );

    const tablesResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-tables",
        tool: "list_tables",
        arguments: {},
      },
    );

    expect(tablesResponse.success).toBe(true);
    const tablesText = tablesResponse.contentItems[0];
    expect(tablesText?.type).toBe("inputText");
    expect(tablesText && "text" in tablesText ? tablesText.text : "").toContain(
      '"name": "RevenueTable"',
    );
    expect(tablesText && "text" in tablesText ? tablesText.text : "").toContain('"columnNames": [');
  });

  it("lists workbook pivots from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setPivotTable("Sheet1", "E2", {
      name: "RevenuePivot",
      source: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      groupBy: ["Revenue"],
      values: [{ sourceColumn: "Margin", summarizeBy: "sum" }],
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-pivots",
        tool: "list_pivots",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"name": "RevenuePivot"');
    expect(text && "text" in text ? text.text : "").toContain('"groupBy": [');
  });

  it("lists workbook data validation rules from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setDataValidation({
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
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
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-data-validations",
        tool: "list_data_validation_rules",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"kind": "list"');
    expect(text && "text" in text ? text.text : "").toContain('"startAddress": "B2"');
  });

  it("includes intersecting data validation metadata in read_range inspection", async () => {
    const engine = await createEngine();
    engine.setDataValidation({
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "list",
        values: ["Draft", "Final"],
      },
      allowBlank: false,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-range-with-validation",
        tool: "read_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"dataValidations": [');
    expect(text && "text" in text ? text.text : "").toContain('"Draft"');
  });

  it("lists workbook conditional formats from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "cellIs",
        operator: "greaterThan",
        values: [10],
      },
      style: {
        fill: { backgroundColor: "#ff0000" },
      },
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-get-conditional-formats",
        tool: "get_conditional_formats",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"id": "cf-1"');
    expect(text && "text" in text ? text.text : "").toContain('"greaterThan"');
  });

  it("includes intersecting conditional format metadata in read_range inspection", async () => {
    const engine = await createEngine();
    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "textContains",
        text: "urgent",
      },
      style: {
        font: { bold: true },
      },
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-range-with-conditional-format",
        tool: "read_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"conditionalFormats": [');
    expect(text && "text" in text ? text.text : "").toContain('"urgent"');
  });

  it("lists workbook protection status from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setSheetProtection({ sheetName: "Sheet1", hideFormulas: true });
    engine.setRangeProtection({
      id: "protect-a1",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B2",
      },
      hideFormulas: true,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-get-protection-status",
        tool: "get_protection_status",
        arguments: { sheetName: "Sheet1" },
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"hideFormulas": true');
    expect(text && "text" in text ? text.text : "").toContain('"protect-a1"');
  });

  it("masks hidden formulas in read_range and inspect_cell outputs", async () => {
    const engine = await createEngine();
    engine.setCellFormula("Sheet1", "A1", "2+2");
    engine.setRangeProtection({
      id: "protect-a1",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A1",
      },
      hideFormulas: true,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const rangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-range-hidden-formula",
        tool: "read_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A1",
        },
      },
    );

    expect(rangeResponse.success).toBe(true);
    const rangeText = rangeResponse.contentItems[0];
    expect(rangeText?.type).toBe("inputText");
    expect(rangeText && "text" in rangeText ? rangeText.text : "").toContain(
      '"rangeProtections": [',
    );
    expect(rangeText && "text" in rangeText ? rangeText.text : "").toContain('"formula": null');

    const inspectResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-inspect-cell-hidden-formula",
        tool: "inspect_cell",
        arguments: {
          sheetName: "Sheet1",
          address: "A1",
        },
      },
    );

    expect(inspectResponse.success).toBe(true);
    const inspectText = inspectResponse.contentItems[0];
    expect(inspectText?.type).toBe("inputText");
    expect(inspectText && "text" in inspectText ? inspectText.text : "").toContain(
      '"formula": null',
    );
  });

  it("lists workbook comments and notes from the authoritative runtime", async () => {
    const engine = await createEngine();
    engine.setCommentThread({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "B2",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    engine.setNote({
      sheetName: "Sheet1",
      address: "C3",
      text: "Manual override",
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-get-comments",
        tool: "get_comments",
        arguments: {},
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"threadId": "thread-1"');
    expect(text && "text" in text ? text.text : "").toContain('"text": "Manual override"');
  });

  it("includes intersecting comment threads and notes in read_range inspection", async () => {
    const engine = await createEngine();
    engine.setCommentThread({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "B2",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    engine.setNote({
      sheetName: "Sheet1",
      address: "C3",
      text: "Manual override",
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "createSheet", name: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-read-range-with-annotations",
        tool: "read_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "C3",
        },
      },
    );

    expect(response.success).toBe(true);
    const text = response.contentItems[0];
    expect(text?.type).toBe("inputText");
    expect(text && "text" in text ? text.text : "").toContain('"commentThreads": [');
    expect(text && "text" in text ? text.text : "").toContain('"notes": [');
    expect(text && "text" in text ? text.text : "").toContain('"Check this total."');
    expect(text && "text" in text ? text.text : "").toContain('"Manual override"');
  });

  it("resolves selectors for read and mutation tools before staging commands", async () => {
    const engine = await createEngine();
    engine.setDefinedName("Inputs", {
      kind: "range-ref",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B1",
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
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const readResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A2",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 3,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-selector-read-range",
        tool: "read_range",
        arguments: {
          selector: {
            kind: "namedRange",
            name: "Inputs",
          },
        },
      },
    );

    expect(readResponse.success).toBe(true);
    const readText = readResponse.contentItems[0];
    expect(readText?.type).toBe("inputText");
    expect(readText && "text" in readText ? readText.text : "").toContain(
      '"displayLabel": "Inputs"',
    );
    expect(readText && "text" in readText ? readText.text : "").toContain('"startAddress": "A1"');

    const formatResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-selector-format-range",
        tool: "format_range",
        arguments: {
          selector: {
            kind: "tableColumn",
            table: "RevenueTable",
            column: "Margin",
          },
          patch: {
            font: {
              bold: true,
            },
          },
        },
      },
    );

    expect(formatResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "formatRange",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B3",
      },
      patch: {
        font: {
          bold: true,
        },
      },
    });
  });

  it("resolves selectors for structural range tools before staging commands", async () => {
    const engine = await createEngine();
    engine.setTable({
      name: "RevenueTable",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Revenue", "Margin"],
      headerRow: true,
      totalsRow: false,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const setFilterResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-selector-set-filter",
        tool: "set_filter",
        arguments: {
          selector: {
            kind: "table",
            table: "RevenueTable",
          },
        },
      },
    );

    expect(setFilterResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "setFilter",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B3",
      },
    });

    const setSortResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-selector-set-sort",
        tool: "set_sort",
        arguments: {
          selector: {
            kind: "tableColumn",
            table: "RevenueTable",
            column: "Margin",
          },
          keys: [{ keyAddress: "B2", direction: "desc" }],
        },
      },
    );

    expect(setSortResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "setSort",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B3",
      },
      keys: [{ keyAddress: "B2", direction: "desc" }],
    });
  });

  it("stages named range, table, and pivot object commands", async () => {
    const engine = await createEngine();
    engine.setTable({
      name: "RevenueTable",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Revenue", "Margin"],
      headerRow: true,
      totalsRow: false,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const namedRangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A2",
            range: {
              startAddress: "A2",
              endAddress: "B3",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-create-named-range",
        tool: "create_named_range",
        arguments: {
          name: "Inputs",
          selector: {
            kind: "currentSelection",
          },
        },
      },
    );

    expect(namedRangeResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "upsertDefinedName",
      name: "Inputs",
      value: {
        kind: "range-ref",
        sheetName: "Sheet1",
        startAddress: "A2",
        endAddress: "B3",
      },
    });

    const tableResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-create-table",
        tool: "create_table",
        arguments: {
          name: "RegionTable",
          selector: {
            kind: "table",
            table: "RevenueTable",
            sheet: "Sheet1",
          },
          headerRow: true,
        },
      },
    );

    expect(tableResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: "upsertTable",
      table: {
        name: "RegionTable",
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "B3",
        columnNames: ["Revenue", "Margin"],
        headerRow: true,
        totalsRow: false,
      },
    });

    const pivotResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-create-pivot",
        tool: "create_pivot_table",
        arguments: {
          name: "RevenuePivot",
          sheetName: "Sheet1",
          address: "E2",
          selector: {
            kind: "table",
            table: "RevenueTable",
            sheet: "Sheet1",
          },
          groupBy: ["Revenue"],
          values: [{ sourceColumn: "Margin", summarizeBy: "sum" }],
        },
      },
    );

    expect(pivotResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "upsertPivotTable",
      pivot: {
        name: "RevenuePivot",
        sheetName: "Sheet1",
        address: "E2",
        source: {
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "B3",
        },
        groupBy: ["Revenue"],
        values: [{ sourceColumn: "Margin", summarizeBy: "sum" }],
        rows: 1,
        cols: 2,
      },
    });
  });

  it("stages selector-aware data validation commands", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const createResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
            range: {
              startAddress: "B2",
              endAddress: "B4",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-create-data-validation",
        tool: "create_data_validation",
        arguments: {
          selector: {
            kind: "currentSelection",
          },
          rule: {
            kind: "list",
            values: ["Draft", "Final"],
          },
          allowBlank: false,
          showDropdown: true,
        },
      },
    );

    expect(createResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "setDataValidation",
      validation: {
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
        rule: {
          kind: "list",
          values: ["Draft", "Final"],
        },
        allowBlank: false,
        showDropdown: true,
      },
    });

    const removeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-remove-data-validation",
        tool: "remove_data_validation",
        arguments: {
          range: {
            sheetName: "Sheet1",
            startAddress: "B2",
            endAddress: "B4",
          },
        },
      },
    );

    expect(removeResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "clearDataValidation",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
    });
  });

  it("stages comment and note commands against single-cell selector targets", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const addCommentResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-add-comment",
        tool: "add_comment",
        arguments: {
          selector: {
            kind: "currentSelection",
          },
          text: "Check this total.",
        },
      },
    );

    expect(addCommentResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "upsertCommentThread",
      thread: {
        threadId: expect.any(String),
        sheetName: "Sheet1",
        address: "B2",
        comments: [{ id: expect.any(String), body: "Check this total." }],
      },
    });

    engine.setCommentThread({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "B2",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });

    const replyCommentResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-reply-comment",
        tool: "reply_comment",
        arguments: {
          range: {
            sheetName: "Sheet1",
            startAddress: "B2",
            endAddress: "B2",
          },
          text: "Looks good.",
        },
      },
    );

    expect(replyCommentResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "upsertCommentThread",
      thread: {
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "B2",
        comments: [
          { id: "comment-1", body: "Check this total." },
          { id: expect.any(String), body: "Looks good." },
        ],
      },
    });

    const addNoteResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-add-note",
        tool: "add_note",
        arguments: {
          range: {
            sheetName: "Sheet1",
            startAddress: "C3",
            endAddress: "C3",
          },
          text: "Manual override",
        },
      },
    );

    expect(addNoteResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "upsertNote",
      note: {
        sheetName: "Sheet1",
        address: "C3",
        text: "Manual override",
      },
    });
  });

  it("stages conditional format commands against selector and id targets", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const addResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
            range: {
              startAddress: "B2",
              endAddress: "B4",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-add-conditional-format",
        tool: "add_conditional_format",
        arguments: {
          selector: {
            kind: "currentSelection",
          },
          rule: {
            kind: "cellIs",
            operator: "greaterThan",
            values: [10],
          },
          style: {
            fill: {
              backgroundColor: "#ff0000",
            },
          },
        },
      },
    );

    expect(addResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "upsertConditionalFormat",
      format: {
        id: expect.any(String),
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
        rule: {
          kind: "cellIs",
          operator: "greaterThan",
          values: [10],
        },
        style: {
          fill: {
            backgroundColor: "#ff0000",
          },
        },
      },
    });

    engine.setConditionalFormat({
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
      rule: {
        kind: "cellIs",
        operator: "greaterThan",
        values: [10],
      },
      style: {
        fill: { backgroundColor: "#ff0000" },
      },
    });

    const removeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-remove-conditional-format",
        tool: "remove_conditional_format",
        arguments: {
          id: "cf-1",
        },
      },
    );

    expect(removeResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "deleteConditionalFormat",
      id: "cf-1",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B4",
      },
    });
  });

  it("stages sheet and range protection commands", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const protectSheetResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "A1",
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-protect-sheet",
        tool: "protect_sheet",
        arguments: {
          hideFormulas: true,
        },
      },
    );

    expect(protectSheetResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "setSheetProtection",
      protection: {
        sheetName: "Sheet1",
        hideFormulas: true,
      },
    });

    const protectRangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: {
          selection: {
            sheetName: "Sheet1",
            address: "B2",
            range: {
              startAddress: "B2",
              endAddress: "B4",
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-protect-range",
        tool: "protect_range",
        arguments: {
          selector: {
            kind: "currentSelection",
          },
          hideFormulas: true,
        },
      },
    );

    expect(protectRangeResponse.success).toBe(true);
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: "upsertRangeProtection",
      protection: {
        id: expect.any(String),
        range: {
          sheetName: "Sheet1",
          startAddress: "B2",
          endAddress: "B4",
        },
        hideFormulas: true,
      },
    });
  });

  it("writes rectangular ranges through renderCommit with normalized formulas", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-2",
        tool: "bilig_write_range",
        arguments: {
          sheetName: "Sheet1",
          startAddress: "C3",
          values: [[1, { formula: "=SUM(A1:A1)" }]],
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "writeRange",
      sheetName: "Sheet1",
      startAddress: "C3",
      values: [[1, { formula: "=SUM(A1:A1)" }]],
    });
    expect(response.contentItems).toEqual([
      expect.objectContaining({
        type: "inputText",
        text: expect.stringContaining('"staged": true'),
      }),
    ]);
    expect(
      response.contentItems[0] && "text" in response.contentItems[0]
        ? response.contentItems[0].text
        : "",
    ).toContain('"bundleId": "bundle-1"');
  });

  it("stages selector-aware formula writes through set_formula", async () => {
    const engine = await createEngine();
    engine.setTable({
      name: "RevenueTable",
      sheetName: "Sheet1",
      startAddress: "A1",
      endAddress: "B3",
      columnNames: ["Revenue", "Margin"],
      headerRow: true,
      totalsRow: false,
    });
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-set-formula",
        tool: "set_formula",
        arguments: {
          selector: {
            kind: "tableColumn",
            table: "RevenueTable",
            column: "Margin",
            sheet: "Sheet1",
          },
          formulas: [["=A2*0.2"], ["=A3*0.25"]],
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "setRangeFormulas",
      range: {
        sheetName: "Sheet1",
        startAddress: "B2",
        endAddress: "B3",
      },
      formulas: [["=A2*0.2"], ["=A3*0.25"]],
    });
  });

  it("stages format commands with normalized number format presets", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-3",
        tool: "bilig_format_range",
        arguments: {
          range: {
            sheetName: "Sheet1",
            startAddress: "A1",
            endAddress: "A2",
          },
          numberFormat: {
            kind: "currency",
            currency: "USD",
          },
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "formatRange",
      range: {
        sheetName: "Sheet1",
        startAddress: "A1",
        endAddress: "A2",
      },
      numberFormat: {
        kind: "currency",
        currency: "USD",
        decimals: 2,
        useGrouping: true,
        negativeStyle: "minus",
        zeroStyle: "zero",
      },
    });
  });

  it("stages row metadata commands for hide and resize operations", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-4",
        tool: "bilig_update_row_metadata",
        arguments: {
          sheetName: "Sheet1",
          startRow: 1,
          count: 2,
          hidden: true,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "updateRowMetadata",
      sheetName: "Sheet1",
      startRow: 1,
      count: 2,
      hidden: true,
    });
  });

  it("stages structural row insertion commands", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-insert-rows",
        tool: "insert_rows",
        arguments: {
          sheetName: "Sheet1",
          start: 1,
          count: 2,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "insertRows",
      sheetName: "Sheet1",
      start: 1,
      count: 2,
    });
  });

  it("stages structural column deletion commands", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-delete-columns",
        tool: "delete_columns",
        arguments: {
          sheetName: "Sheet1",
          start: 0,
          count: 1,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "deleteColumns",
      sheetName: "Sheet1",
      start: 0,
      count: 1,
    });
  });

  it.each([
    {
      name: "delete_sheet",
      request: { name: "Imports" },
      expected: { kind: "deleteSheet", name: "Imports" },
    },
    {
      name: "set_freeze_pane",
      request: { sheetName: "Sheet1", rows: 1, cols: 2 },
      expected: { kind: "setFreezePane", sheetName: "Sheet1", rows: 1, cols: 2 },
    },
    {
      name: "set_filter",
      request: {
        range: { sheetName: "Sheet1", startAddress: "D4", endAddress: "A1" },
      },
      expected: {
        kind: "setFilter",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "D4" },
      },
    },
    {
      name: "clear_filter",
      request: {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "D4" },
      },
      expected: {
        kind: "clearFilter",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "D4" },
      },
    },
    {
      name: "set_sort",
      request: {
        range: { sheetName: "Sheet1", startAddress: "B3", endAddress: "A1" },
        keys: [{ keyAddress: "B1", direction: "desc" as const }],
      },
      expected: {
        kind: "setSort",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
        keys: [{ keyAddress: "B1", direction: "desc" as const }],
      },
    },
    {
      name: "clear_sort",
      request: {
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
      expected: {
        kind: "clearSort",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "B3" },
      },
    },
  ])("stages $name commands", async ({ name, request, expected }) => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: `call-${name}`,
        tool: name,
        arguments: request,
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith(expected);
  });

  it("starts built-in durable workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeWorkbook" as const,
      title: "Summarize Workbook",
      summary: "Summarized workbook structure across 2 sheets.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-workbook",
          label: "Inspect workbook structure",
          status: "completed" as const,
          summary: "Read durable workbook structure across 2 sheets.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "draft-summary",
          label: "Draft summary artifact",
          status: "completed" as const,
          summary: "Prepared the durable workbook summary artifact for the thread.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Workbook Summary",
        text: "## Workbook Summary",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "summarizeWorkbook",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "summarizeWorkbook",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain('"runId": "wf-1"');
    expect(output && "text" in output ? output.text : "").toContain('"title": "Workbook Summary"');
  });

  it("starts structural create-sheet workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-create-sheet-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "createSheet" as const,
      title: "Create Sheet",
      summary: "Staged a structural change set to create Forecast.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "plan-sheet-create",
          label: "Plan sheet creation",
          status: "completed" as const,
          summary: "Prepared the semantic sheet-creation command for Forecast.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "stage-structural-preview",
          label: "Stage structural preview",
          status: "completed" as const,
          summary: "Staged the structural change set in the thread rail.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Create Sheet Preview",
        text: "## Create Sheet Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-create-sheet-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "createSheet",
          name: "Forecast",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "createSheet",
      name: "Forecast",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "createSheet"',
    );
  });

  it("starts query-driven workbook search workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-search-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "searchWorkbookQuery" as const,
      title: "Search Workbook",
      summary: 'Found 2 workbook matches for "revenue".',
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "search-workbook",
          label: "Search workbook",
          status: "completed" as const,
          summary:
            'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 2 matches.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: "draft-search-report",
          label: "Draft search report",
          status: "completed" as const,
          summary: "Prepared the durable workbook search report for the thread.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Workbook Search",
        text: "## Workbook Search",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-search-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "searchWorkbookQuery",
          query: "revenue",
          limit: 10,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "searchWorkbookQuery",
      query: "revenue",
      limit: 10,
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain('"runId": "wf-search-1"');
    expect(output && "text" in output ? output.text : "").toContain('"title": "Search Workbook"');
  });

  it("starts current-sheet summary workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-sheet-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeCurrentSheet" as const,
      title: "Summarize Current Sheet",
      summary: "Summarized Revenue with 24 populated cells and 1 table.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-current-sheet",
          label: "Inspect current sheet",
          status: "completed" as const,
          summary: "Read durable metadata for Revenue, including used range and tables.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "draft-sheet-summary",
          label: "Draft current sheet summary",
          status: "completed" as const,
          summary: "Prepared the durable current-sheet summary artifact for the thread.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Current Sheet Summary",
        text: "## Current Sheet Summary",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-sheet-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "summarizeCurrentSheet",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "summarizeCurrentSheet",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain('"runId": "wf-sheet-1"');
    expect(output && "text" in output ? output.text : "").toContain(
      '"title": "Current Sheet Summary"',
    );
  });

  it("starts sheet-scoped formula issue workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-formula-sheet-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "findFormulaIssues" as const,
      title: "Find Formula Issues",
      summary: "Found 2 formula issues on Sheet1 across 3 scanned formula cells.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "scan-formula-cells",
          label: "Scan formula cells",
          status: "completed" as const,
          summary: "Scanned 3 formula cells on Sheet1 and found 2 issues.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "draft-issue-report",
          label: "Draft issue report",
          status: "completed" as const,
          summary: "Prepared the durable formula issue report for the thread.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Formula Issues",
        text: "## Formula Issues",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-formula-sheet-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "findFormulaIssues",
          sheetName: "Sheet1",
          limit: 25,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "findFormulaIssues",
      sheetName: "Sheet1",
      limit: 25,
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"runId": "wf-formula-sheet-1"',
    );
  });

  it("starts highlight-formula workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-formula-highlight-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "highlightFormulaIssues" as const,
      title: "Highlight Formula Issues",
      summary: "Staged highlight formatting for 2 formula issues on Sheet1.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "scan-formula-cells",
          label: "Scan formula cells",
          status: "completed" as const,
          summary: "Scanned 3 formula cells on Sheet1 and found 2 issues.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "stage-issue-highlights",
          label: "Stage issue highlights",
          status: "completed" as const,
          summary:
            "Prepared 2 semantic formatting commands to highlight the detected formula issues.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Formula Issue Highlights",
        text: "## Highlighted Formula Issues",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-highlight-formula-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "highlightFormulaIssues",
          sheetName: "Sheet1",
          limit: 25,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "highlightFormulaIssues",
      sheetName: "Sheet1",
      limit: 25,
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "highlightFormulaIssues"',
    );
  });

  it("starts repair-formula workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-formula-repair-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "repairFormulaIssues" as const,
      title: "Repair Formula Issues",
      summary: "Staged 1 formula repair on Sheet1 from nearby healthy formulas.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "scan-formula-cells",
          label: "Scan formula cells",
          status: "completed" as const,
          summary: "Scanned 2 formula cells on Sheet1 and found 1 issue.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "stage-formula-repairs",
          label: "Stage formula repairs",
          status: "completed" as const,
          summary: "Prepared 1 semantic write command for the repair change set.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Formula Repair Preview",
        text: "## Formula Repair Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-repair-formula-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "repairFormulaIssues",
          sheetName: "Sheet1",
          limit: 25,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "repairFormulaIssues",
      sheetName: "Sheet1",
      limit: 25,
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "repairFormulaIssues"',
    );
  });

  it("starts outlier-highlight workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-outlier-highlight-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "highlightCurrentSheetOutliers" as const,
      title: "Highlight Current Sheet Outliers",
      summary: "Staged outlier highlights for 2 cells across 1 numeric column on Revenue.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-numeric-columns",
          label: "Inspect numeric columns",
          status: "completed" as const,
          summary: "Loaded numeric cells and header labels from Revenue.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Current Sheet Outlier Highlights",
        text: "## Highlighted Numeric Outliers",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-outlier-highlight-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "highlightCurrentSheetOutliers",
          sheetName: "Revenue",
          limit: 10,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "highlightCurrentSheetOutliers",
      sheetName: "Revenue",
      limit: 10,
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "highlightCurrentSheetOutliers"',
    );
  });

  it("starts header-normalization workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-header-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "normalizeCurrentSheetHeaders" as const,
      title: "Normalize Current Sheet Headers",
      summary: "Staged normalized headers for 2 cells on Imports.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-header-row",
          label: "Inspect header row",
          status: "completed" as const,
          summary: "Loaded the used range and current header row from Imports.",
          updatedAtUnixMs: 1,
        },
        {
          stepId: "stage-header-normalization",
          label: "Stage header normalization",
          status: "completed" as const,
          summary: "Prepared the semantic write preview that normalizes 2 header cells.",
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Header Normalization Preview",
        text: "## Header Normalization Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-header-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "normalizeCurrentSheetHeaders",
          sheetName: "Imports",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "normalizeCurrentSheetHeaders",
      sheetName: "Imports",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "normalizeCurrentSheetHeaders"',
    );
  });

  it("starts number-format-normalization workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-number-format-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "normalizeCurrentSheetNumberFormats" as const,
      title: "Normalize Current Sheet Number Formats",
      summary: "Staged normalized number formats for 3 columns on Imports.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-number-columns",
          label: "Inspect numeric columns",
          status: "completed" as const,
          summary: "Loaded numeric cells and header labels from Imports.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Number Format Normalization Preview",
        text: "## Number Format Normalization Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-number-format-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "normalizeCurrentSheetNumberFormats",
          sheetName: "Imports",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "normalizeCurrentSheetNumberFormats",
      sheetName: "Imports",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "normalizeCurrentSheetNumberFormats"',
    );
  });

  it("starts whitespace-normalization workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-whitespace-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "normalizeCurrentSheetWhitespace" as const,
      title: "Normalize Current Sheet Whitespace",
      summary: "Staged normalized whitespace for 3 text cells on Imports.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-text-cells",
          label: "Inspect text cells",
          status: "completed" as const,
          summary: "Loaded the used range and string cells from Imports.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Whitespace Normalization Preview",
        text: "## Whitespace Normalization Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-whitespace-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "normalizeCurrentSheetWhitespace",
          sheetName: "Imports",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "normalizeCurrentSheetWhitespace",
      sheetName: "Imports",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "normalizeCurrentSheetWhitespace"',
    );
  });

  it("starts formula fill-down workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-fill-formulas-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "fillCurrentSheetFormulasDown" as const,
      title: "Fill Current Sheet Formulas Down",
      summary: "Staged formula fill-down for 1 column on Imports.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-formula-columns",
          label: "Inspect formula columns",
          status: "completed" as const,
          summary: "Loaded formula cells and blank fill gaps from Imports.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Formula Fill-Down Preview",
        text: "## Formula Fill-Down Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-fill-formulas-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "fillCurrentSheetFormulasDown",
          sheetName: "Imports",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "fillCurrentSheetFormulasDown",
      sheetName: "Imports",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "fillCurrentSheetFormulasDown"',
    );
  });

  it("starts header-style workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-style-headers-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "styleCurrentSheetHeaders" as const,
      title: "Style Current Sheet Headers",
      summary: "Staged a consistent header style preview for Imports.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-header-row",
          label: "Inspect header row",
          status: "completed" as const,
          summary: "Loaded the used range and header row from Imports.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Header Style Preview",
        text: "## Header Style Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-style-headers-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "styleCurrentSheetHeaders",
          sheetName: "Imports",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "styleCurrentSheetHeaders",
      sheetName: "Imports",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "styleCurrentSheetHeaders"',
    );
  });

  it("starts current-sheet review-tab workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-review-tab-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "createCurrentSheetReviewTab" as const,
      title: "Create Current Sheet Review Tab",
      summary: "Staged a review-tab preview for Revenue into Revenue Review.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-source-sheet",
          label: "Inspect source sheet",
          status: "completed" as const,
          summary: "Loaded the used range from Revenue.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Current Sheet Review Tab Preview",
        text: "## Current Sheet Review Tab Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-review-tab-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "createCurrentSheetReviewTab",
          sheetName: "Revenue",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "createCurrentSheetReviewTab",
      sheetName: "Revenue",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "createCurrentSheetReviewTab"',
    );
  });

  it("starts current-sheet rollup workflows from the semantic tool surface", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const startWorkflow = vi.fn(async () => ({
      runId: "wf-rollup-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "createCurrentSheetRollup" as const,
      title: "Create Current Sheet Rollup",
      summary: "Staged a rollup preview for Revenue into Revenue Rollup.",
      status: "completed" as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-source-sheet",
          label: "Inspect source sheet",
          status: "completed" as const,
          summary: "Loaded the used range and numeric columns from Revenue.",
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: "markdown" as const,
        title: "Current Sheet Rollup Preview",
        text: "## Current Sheet Rollup Preview",
      },
    }));

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
          createBundle(command),
        ),
        startWorkflow,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-workflow-rollup-1",
        tool: "bilig_start_workflow",
        arguments: {
          workflowTemplate: "createCurrentSheetRollup",
          sheetName: "Revenue",
        },
      },
    );

    expect(response.success).toBe(true);
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: "createCurrentSheetRollup",
      sheetName: "Revenue",
    });
    const output = response.contentItems.find((item) => item.type === "inputText");
    expect(output?.type).toBe("inputText");
    expect(output && "text" in output ? output.text : "").toContain(
      '"workflowTemplate": "createCurrentSheetRollup"',
    );
  });

  it("stages column metadata commands for resize operations", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle["commands"][number]) =>
      createBundle(command),
    );

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-5",
        tool: "bilig_update_column_metadata",
        arguments: {
          sheetName: "Sheet1",
          startCol: 0,
          count: 2,
          width: 120,
        },
      },
    );

    expect(response.success).toBe(true);
    expect(stageCommand).toHaveBeenCalledWith({
      kind: "updateColumnMetadata",
      sheetName: "Sheet1",
      startCol: 0,
      count: 2,
      width: 120,
    });
  });
});
