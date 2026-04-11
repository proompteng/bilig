import { SpreadsheetEngine } from "@bilig/core";
import { describe, expect, it } from "vitest";
import { buildWorkbookSourceProjectionFromEngine } from "../zero/projection.js";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";
import {
  describeWorkbookAgentWorkflowTemplate,
  executeWorkbookAgentWorkflow,
} from "./workbook-agent-workflows.js";

async function createWorkbookRuntime(): Promise<WorkbookRuntime> {
  const engine = new SpreadsheetEngine({
    workbookName: "doc-1",
    replicaId: "server:test",
  });
  await engine.ready();
  engine.createSheet("Sheet1");
  engine.setCellValue("Sheet1", "A1", 42);
  engine.setCellValue("Sheet1", "A2", "Gross Margin");
  engine.setCellFormula("Sheet1", "B2", "SUM(A1:A1)");
  return {
    documentId: "doc-1",
    engine,
    projection: buildWorkbookSourceProjectionFromEngine("doc-1", engine, {
      revision: 1,
      calculatedRevision: 1,
      ownerUserId: "alex@example.com",
      updatedBy: "alex@example.com",
      updatedAt: "2026-04-10T00:00:00.000Z",
    }),
    headRevision: 1,
    calculatedRevision: 1,
    ownerUserId: "alex@example.com",
  };
}

function createZeroSyncStub(input?: {
  onInspectWorkbook?: () => void;
  createRuntime?: () => Promise<WorkbookRuntime>;
}): ZeroSyncService {
  return {
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
      input?.onInspectWorkbook?.();
      if (input?.createRuntime) {
        return await task(await input.createRuntime());
      }
      throw new Error("inspectWorkbook should not be called");
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
}

describe("workbook agent workflows", () => {
  it("describes structural workflow templates through the structural metadata path", () => {
    expect(
      describeWorkbookAgentWorkflowTemplate("createSheet", {
        name: "Ops Review",
      }),
    ).toEqual({
      title: "Create Sheet",
      runningSummary: "Preparing a structural preview bundle to create Ops Review.",
    });
  });

  it("executes structural workflow templates without durable workbook inspection", async () => {
    let inspectedWorkbook = false;
    const result = await executeWorkbookAgentWorkflow({
      documentId: "doc-1",
      zeroSyncService: createZeroSyncStub({
        onInspectWorkbook: () => {
          inspectedWorkbook = true;
        },
      }),
      workflowTemplate: "hideCurrentRow",
      context: {
        selection: {
          sheetName: "Sheet1",
          address: "B3",
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      },
    });

    expect(inspectedWorkbook).toBe(false);
    expect(result.title).toBe("Hide Current Row");
    expect(result.commands).toEqual([
      {
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        startRow: 2,
        count: 1,
        hidden: true,
      },
    ]);
  });

  it("executes summarize workbook through the durable inspection path", async () => {
    let inspectedWorkbook = false;
    const result = await executeWorkbookAgentWorkflow({
      documentId: "doc-1",
      zeroSyncService: createZeroSyncStub({
        onInspectWorkbook: () => {
          inspectedWorkbook = true;
        },
        createRuntime: createWorkbookRuntime,
      }),
      workflowTemplate: "summarizeWorkbook",
    });

    expect(inspectedWorkbook).toBe(true);
    expect(result.title).toBe("Summarize Workbook");
    expect(result.summary).toContain("Summarized workbook structure across 1 sheet");
  });
});
