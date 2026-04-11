import { describe, expect, it } from "vitest";
import type { ZeroSyncService } from "../zero/service.js";
import {
  describeWorkbookAgentWorkflowTemplate,
  executeWorkbookAgentWorkflow,
} from "./workbook-agent-workflows.js";

function createZeroSyncStub(input?: { onInspectWorkbook?: () => void }): ZeroSyncService {
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
    async inspectWorkbook() {
      input?.onInspectWorkbook?.();
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
});
