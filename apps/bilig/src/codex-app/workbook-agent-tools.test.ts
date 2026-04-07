import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import { handleWorkbookAgentToolCall } from "./workbook-agent-tools.js";

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
    async listWorkbookAgentRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {
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

describe("workbook agent tools", () => {
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
        tool: "bilig.read_selection",
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
        tool: "bilig.read_visible_range",
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
        tool: "bilig.inspect_cell",
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
        tool: "bilig.find_formula_issues",
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
        tool: "bilig.search_workbook",
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
        tool: "bilig.trace_dependencies",
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
        tool: "bilig.read_range",
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
        tool: "bilig.write_range",
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
        tool: "bilig.format_range",
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
});
