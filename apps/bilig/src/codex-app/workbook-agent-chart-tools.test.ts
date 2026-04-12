import { SpreadsheetEngine } from "@bilig/core";
import {
  createWorkbookAgentCommandBundle,
  WORKBOOK_AGENT_TOOL_NAMES,
  type CodexDynamicToolCallResult,
  type WorkbookAgentCommand,
} from "@bilig/agent-api";
import { describe, expect, it, vi } from "vitest";
import type { AuthoritativeWorkbookEventBatch } from "@bilig/zero-sync";
import { z } from "zod";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import { handleWorkbookAgentToolCall } from "./workbook-agent-tools.js";

async function createEngine(): Promise<SpreadsheetEngine> {
  const engine = new SpreadsheetEngine({
    workbookName: "doc-1",
    replicaId: "server:chart-test",
  });
  await engine.ready();
  engine.createSheet("Data");
  engine.createSheet("Dashboard");
  engine.setRangeValues({ sheetName: "Data", startAddress: "A1", endAddress: "B4" }, [
    ["Month", "Revenue"],
    ["Jan", 10],
    ["Feb", 15],
    ["Mar", 9],
  ]);
  engine.setChart({
    id: "Revenue Chart",
    sheetName: "Dashboard",
    address: "B2",
    source: { sheetName: "Data", startAddress: "A1", endAddress: "B4" },
    chartType: "column",
    rows: 12,
    cols: 8,
    title: "Revenue",
  });
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
          updatedAt: "2026-04-12T12:00:00.000Z",
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
      throw new Error("not used");
    },
    async listWorkbookChanges() {
      return [];
    },
    async listWorkbookAgentRuns() {
      return [];
    },
    async appendWorkbookAgentRun() {
      throw new Error("not used");
    },
    async listWorkbookAgentThreadRuns() {
      return [];
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
      return {
        afterRevision: 1,
        headRevision: 1,
        calculatedRevision: 1,
        events: [],
      } satisfies AuthoritativeWorkbookEventBatch;
    },
  };
  return { zeroSyncService };
}

function createBundle(command: WorkbookAgentCommand) {
  return createWorkbookAgentCommandBundle({
    documentId: "doc-1",
    threadId: "thr-1",
    turnId: "turn-1",
    goalText: "chart test",
    baseRevision: 1,
    now: 1,
    context: null,
    commands: [command],
  });
}

function parsePayload(result: CodexDynamicToolCallResult): unknown {
  expect(result.success).toBe(true);
  const item = result.contentItems[0];
  expect(item?.type).toBe("inputText");
  return JSON.parse(item && "text" in item ? item.text : "");
}

const listChartsPayloadSchema = z.object({
  chartCount: z.number(),
  charts: z.array(
    z.object({
      id: z.string(),
      sheetName: z.string(),
      address: z.string(),
      chartType: z.string(),
    }),
  ),
});

const stagedChartPayloadSchema = z.object({
  staged: z.boolean(),
  bundleId: z.string(),
  affectedRanges: z.array(
    z.object({
      sheetName: z.string(),
      startAddress: z.string(),
      endAddress: z.string(),
      role: z.string(),
    }),
  ),
});

describe("workbook agent chart tools", () => {
  it("lists workbook charts from the authoritative runtime", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);

    const result = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: { userID: "alex@example.com", roles: ["editor"] },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: "deleteChart", id: "unused" })),
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-list-charts",
        tool: WORKBOOK_AGENT_TOOL_NAMES.listCharts,
        arguments: {},
      },
    );

    const payload = listChartsPayloadSchema.parse(parsePayload(result));
    expect(payload.chartCount).toBe(1);
    expect(payload.charts).toContainEqual(
      expect.objectContaining({
        id: "Revenue Chart",
        sheetName: "Dashboard",
        address: "B2",
        chartType: "column",
      }),
    );
  });

  it("stages create-chart and delete-chart commands", async () => {
    const engine = await createEngine();
    const { zeroSyncService } = createZeroSyncHarness(engine);
    const stageCommand = vi.fn(async (command: WorkbookAgentCommand) => createBundle(command));

    const createResult = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: { userID: "alex@example.com", roles: ["editor"] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-create-chart",
        tool: WORKBOOK_AGENT_TOOL_NAMES.createChart,
        arguments: {
          id: "Margin Chart",
          sheetName: "Dashboard",
          address: "H2",
          range: {
            sheetName: "Data",
            startAddress: "A1",
            endAddress: "B4",
          },
          chartType: "line",
          title: "Margin",
        },
      },
    );
    const deleteResult = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: { userID: "alex@example.com", roles: ["editor"] },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: "thr-1",
        turnId: "turn-1",
        callId: "call-delete-chart",
        tool: WORKBOOK_AGENT_TOOL_NAMES.deleteChart,
        arguments: {
          id: "Revenue Chart",
        },
      },
    );

    expect(stageCommand).toHaveBeenNthCalledWith(
      1,
      expect.objectContaining({
        kind: "upsertChart",
        chart: expect.objectContaining({
          id: "Margin Chart",
          sheetName: "Dashboard",
          address: "H2",
          chartType: "line",
          source: {
            sheetName: "Data",
            startAddress: "A1",
            endAddress: "B4",
          },
        }),
      }),
    );
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: "deleteChart",
      id: "Revenue Chart",
    });

    expect(stagedChartPayloadSchema.parse(parsePayload(createResult))).toEqual(
      expect.objectContaining({
        staged: true,
        affectedRanges: expect.arrayContaining([
          expect.objectContaining({
            sheetName: "Data",
            startAddress: "A1",
            endAddress: "B4",
            role: "source",
          }),
          expect.objectContaining({
            sheetName: "Dashboard",
            startAddress: "H2",
            role: "target",
          }),
        ]),
      }),
    );
    expect(stagedChartPayloadSchema.parse(parsePayload(deleteResult))).toEqual(
      expect.objectContaining({
        staged: true,
      }),
    );
  });
});
