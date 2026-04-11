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
    approvalMode: "auto",
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
            address: "A1",
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
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"startAddress": "A1"');
    expect(textItem && "text" in textItem ? textItem.text : "").toContain('"endAddress": "A1"');
  });

  it("reads the visible viewport through the attached browser context", async () => {
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
        summary: "Applied preview bundle at revision r12",
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
    expect(text).toContain('"summary": "Applied preview bundle at revision r12"');
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
      summary: "Staged a structural preview bundle to create Forecast.",
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
          summary: "Staged the structural preview bundle in the thread rail.",
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
          summary: 'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 2 matches.',
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
