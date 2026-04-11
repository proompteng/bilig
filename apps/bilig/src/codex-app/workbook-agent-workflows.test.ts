import { describe, expect, it } from "vitest";
import type { ZeroSyncService } from "../zero/service.js";
import type { WorkbookStructureSummary } from "./workbook-agent-comprehension.js";
import {
  describeWorkbookAgentWorkflowTemplate,
  executeWorkbookAgentWorkflow,
} from "./workbook-agent-workflows.js";

function createZeroSyncStub(input?: {
  onInspectWorkbook?: () => void;
  inspectWorkbookResult?: WorkbookStructureSummary;
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
    async inspectWorkbook() {
      input?.onInspectWorkbook?.();
      if (input?.inspectWorkbookResult) {
        return input.inspectWorkbookResult;
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
        inspectWorkbookResult: {
          summary: {
            sheetCount: 1,
            totalCellCount: 3,
            totalFormulaCellCount: 1,
            tableCount: 0,
            pivotCount: 0,
            spillCount: 0,
            filterCount: 0,
            sortCount: 0,
            hiddenRowIndexCount: 0,
            hiddenColumnIndexCount: 0,
          },
          sheets: [
            {
              name: "Sheet1",
              order: 0,
              cellCount: 3,
              formulaCellCount: 1,
              usedRange: {
                startAddress: "A1",
                endAddress: "B2",
              },
              freezePane: null,
              filterCount: 0,
              sortCount: 0,
              tableCount: 0,
              pivotCount: 0,
              spillCount: 0,
              rowMetadata: {
                regionCount: 0,
                hiddenIndexCount: 0,
                explicitSizeIndexCount: 0,
              },
              columnMetadata: {
                regionCount: 0,
                hiddenIndexCount: 0,
                explicitSizeIndexCount: 0,
              },
              tables: [],
              pivots: [],
              spills: [],
            },
          ],
        },
      }),
      workflowTemplate: "summarizeWorkbook",
    });

    expect(inspectedWorkbook).toBe(true);
    expect(result.title).toBe("Summarize Workbook");
    expect(result.summary).toContain("Summarized workbook structure across 1 sheet");
  });
});
