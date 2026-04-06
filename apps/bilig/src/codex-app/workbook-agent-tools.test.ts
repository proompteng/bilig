import { describe, expect, it, vi } from "vitest";
import { SpreadsheetEngine } from "@bilig/core";
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
  engine.setCellFormula("Sheet1", "B1", "SUM(A1:A1)");
  return engine;
}

function createZeroSyncHarness(engine: SpreadsheetEngine) {
  const applyServerMutator = vi.fn(async () => undefined);
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
    applyServerMutator,
    async loadAuthoritativeEvents() {
      throw new Error("not used");
    },
  };
  return {
    zeroSyncService,
    applyServerMutator,
  };
}

describe("workbook agent tools", () => {
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
    const { zeroSyncService, applyServerMutator } = createZeroSyncHarness(engine);

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: "doc-1",
        session: {
          userID: "alex@example.com",
          roles: ["editor"],
        },
        uiContext: null,
        zeroSyncService,
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
    expect(applyServerMutator).toHaveBeenCalledWith(
      "workbook.renderCommit",
      expect.objectContaining({
        documentId: "doc-1",
        ops: [
          {
            kind: "setCellValue",
            sheetName: "Sheet1",
            address: "C3",
            value: 1,
          },
          {
            kind: "setCellFormula",
            sheetName: "Sheet1",
            address: "D3",
            formula: "SUM(A1:A1)",
          },
        ],
      }),
      expect.objectContaining({
        userID: "alex@example.com",
      }),
    );
  });
});
