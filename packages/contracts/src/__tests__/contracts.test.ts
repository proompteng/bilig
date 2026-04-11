import { describe, expect, it } from "vitest";
import {
  decodeUnknownSync,
  DocumentStateSummarySchema,
  ErrorEnvelopeSchema,
  RuntimeSessionSchema,
  WorkbookAgentTimelineEntrySchema,
  WorkbookAgentThreadSummarySchema,
  WorkbookAgentWorkflowRunSchema,
} from "../index.js";

describe("@bilig/contracts", () => {
  it("decodes a v2 runtime session payload", () => {
    const decoded = decodeUnknownSync(RuntimeSessionSchema, {
      authToken: "user-123",
      userId: "user-123",
      roles: ["editor"],
      isAuthenticated: true,
      authSource: "header",
    });

    expect(decoded.authToken).toBe("user-123");
    expect(decoded.authSource).toBe("header");
  });

  it("decodes a v2 document state payload", () => {
    const decoded = decodeUnknownSync(DocumentStateSummarySchema, {
      documentId: "book-1",
      cursor: 4,
      owner: null,
      sessions: ["browser:1"],
      latestSnapshotCursor: 3,
    });

    expect(decoded.documentId).toBe("book-1");
    expect(decoded.latestSnapshotCursor).toBe(3);
  });

  it("decodes a v2 error envelope", () => {
    const decoded = decodeUnknownSync(ErrorEnvelopeSchema, {
      error: "TEST_FAILURE",
      message: "boom",
      retryable: false,
    });

    expect(decoded.retryable).toBe(false);
  });

  it("decodes workbook agent thread summaries with owner metadata", () => {
    const decoded = decodeUnknownSync(WorkbookAgentThreadSummarySchema, {
      threadId: "thr-1",
      scope: "shared",
      ownerUserId: "alex@example.com",
      updatedAtUnixMs: 42,
      entryCount: 3,
      hasPendingBundle: true,
      latestEntryText: "Preview bundle staged",
    });

    expect(decoded.ownerUserId).toBe("alex@example.com");
    expect(decoded.hasPendingBundle).toBe(true);
  });

  it("decodes workbook agent timeline entries with citations", () => {
    const decoded = decodeUnknownSync(WorkbookAgentTimelineEntrySchema, {
      id: "system-1",
      kind: "system",
      turnId: "turn-1",
      text: "Applied preview bundle at revision r7",
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [
        {
          kind: "range",
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A3",
          role: "target",
        },
        {
          kind: "revision",
          revision: 7,
        },
      ],
    });

    expect(decoded.citations).toEqual([
      expect.objectContaining({
        kind: "range",
        sheetName: "Sheet1",
      }),
      expect.objectContaining({
        kind: "revision",
        revision: 7,
      }),
    ]);
  });

  it("decodes workbook agent workflow runs with markdown artifacts", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeWorkbook",
      title: "Summarize Workbook",
      summary: "Summarized workbook structure across 3 sheets.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-workbook",
          label: "Inspect workbook structure",
          status: "completed",
          summary: "Read durable workbook structure across 3 sheets.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "draft-summary",
          label: "Draft summary artifact",
          status: "completed",
          summary: "Prepared the durable workbook summary artifact for the thread.",
          updatedAtUnixMs: 120,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Workbook Summary",
        text: "## Summary",
      },
    });

    expect(decoded.workflowTemplate).toBe("summarizeWorkbook");
    expect(decoded.steps).toHaveLength(2);
    expect(decoded.steps[0]?.status).toBe("completed");
    expect(decoded.artifact).toEqual(
      expect.objectContaining({
        kind: "markdown",
        title: "Workbook Summary",
      }),
    );
  });

  it("accepts explain-selection workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-2",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "explainSelectionCell",
      title: "Explain Current Cell",
      summary: "Explained Sheet1!B2, including direct precedents and dependents.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-selection",
          label: "Inspect current selection",
          status: "completed",
          summary: "Loaded workbook context for Sheet1!B2.",
          updatedAtUnixMs: 110,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Current Cell",
        text: "## Current Cell",
      },
    });

    expect(decoded.workflowTemplate).toBe("explainSelectionCell");
    expect(decoded.artifact?.title).toBe("Current Cell");
  });

  it("accepts current-sheet summary workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-sheet-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeCurrentSheet",
      title: "Summarize Current Sheet",
      summary: "Summarized Revenue with 24 populated cells and 1 table.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-current-sheet",
          label: "Inspect current sheet",
          status: "completed",
          summary: "Read durable metadata for Revenue, including used range and tables.",
          updatedAtUnixMs: 110,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Current Sheet Summary",
        text: "## Current Sheet Summary",
      },
    });

    expect(decoded.workflowTemplate).toBe("summarizeCurrentSheet");
    expect(decoded.artifact?.title).toBe("Current Sheet Summary");
  });

  it("accepts search-query workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-3",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "searchWorkbookQuery",
      title: "Search Workbook",
      summary: 'Found 2 workbook matches for "revenue".',
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "search-workbook",
          label: "Search workbook",
          status: "completed",
          summary:
            'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 2 matches.',
          updatedAtUnixMs: 110,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Workbook Search",
        text: "## Workbook Search",
      },
    });

    expect(decoded.workflowTemplate).toBe("searchWorkbookQuery");
    expect(decoded.artifact?.title).toBe("Workbook Search");
  });

  it("accepts highlight-formula workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-highlight-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "highlightFormulaIssues",
      title: "Highlight Formula Issues",
      summary: "Staged highlight formatting for 2 formula issues on Sheet1.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "scan-formula-cells",
          label: "Scan formula cells",
          status: "completed",
          summary: "Scanned 3 formula cells on Sheet1 and found 2 issues.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-issue-highlights",
          label: "Stage issue highlights",
          status: "completed",
          summary:
            "Prepared 2 semantic formatting commands to highlight the detected formula issues.",
          updatedAtUnixMs: 115,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Formula Issue Highlights",
        text: "## Highlighted Formula Issues",
      },
    });

    expect(decoded.workflowTemplate).toBe("highlightFormulaIssues");
    expect(decoded.artifact?.title).toBe("Formula Issue Highlights");
  });

  it("accepts outlier-highlight workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-outlier-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "highlightCurrentSheetOutliers",
      title: "Highlight Current Sheet Outliers",
      summary: "Staged outlier highlights for 2 cells across 1 numeric column on Revenue.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-numeric-columns",
          label: "Inspect numeric columns",
          status: "completed",
          summary: "Loaded numeric cells and header labels from Revenue.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-outlier-highlights",
          label: "Stage outlier highlights",
          status: "completed",
          summary:
            "Prepared 2 semantic formatting commands to highlight numeric outliers on Revenue.",
          updatedAtUnixMs: 115,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Current Sheet Outlier Highlights",
        text: "## Highlighted Numeric Outliers",
      },
    });

    expect(decoded.workflowTemplate).toBe("highlightCurrentSheetOutliers");
    expect(decoded.artifact?.title).toBe("Current Sheet Outlier Highlights");
  });

  it("accepts header-normalization workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-header-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "normalizeCurrentSheetHeaders",
      title: "Normalize Current Sheet Headers",
      summary: "Staged normalized headers for 2 cells on Imports.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-header-row",
          label: "Inspect header row",
          status: "completed",
          summary: "Loaded the used range and current header row from Imports.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-header-normalization",
          label: "Stage header normalization",
          status: "completed",
          summary: "Prepared the semantic write preview that normalizes 2 header cells.",
          updatedAtUnixMs: 115,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Header Normalization Preview",
        text: "## Header Normalization Preview",
      },
    });

    expect(decoded.workflowTemplate).toBe("normalizeCurrentSheetHeaders");
    expect(decoded.artifact?.title).toBe("Header Normalization Preview");
  });

  it("accepts number-format-normalization workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-number-format-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "normalizeCurrentSheetNumberFormats",
      title: "Normalize Current Sheet Number Formats",
      summary: "Staged normalized number formats for 3 columns on Imports.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-number-columns",
          label: "Inspect numeric columns",
          status: "completed",
          summary: "Loaded numeric cells and header labels from Imports.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-number-formats",
          label: "Stage number formats",
          status: "completed",
          summary: "Prepared semantic number-format previews for 3 numeric columns.",
          updatedAtUnixMs: 115,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Number Format Normalization Preview",
        text: "## Number Format Normalization Preview",
      },
    });

    expect(decoded.workflowTemplate).toBe("normalizeCurrentSheetNumberFormats");
    expect(decoded.artifact?.title).toBe("Number Format Normalization Preview");
  });

  it("accepts rollup workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-rollup-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "createCurrentSheetRollup",
      title: "Create Current Sheet Rollup",
      summary: "Staged a rollup preview for Revenue into Revenue Rollup.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "inspect-source-sheet",
          label: "Inspect source sheet",
          status: "completed",
          summary: "Loaded the used range and numeric columns from Revenue.",
          updatedAtUnixMs: 110,
        },
        {
          stepId: "stage-rollup-preview",
          label: "Stage rollup preview",
          status: "completed",
          summary: "Prepared the semantic preview that creates Revenue Rollup.",
          updatedAtUnixMs: 115,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Current Sheet Rollup Preview",
        text: "## Current Sheet Rollup Preview",
      },
    });

    expect(decoded.workflowTemplate).toBe("createCurrentSheetRollup");
    expect(decoded.artifact?.title).toBe("Current Sheet Rollup Preview");
  });

  it("accepts structural workflow templates in durable runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-structural-1",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "createSheet",
      title: "Create Sheet",
      summary: "Staged a structural preview bundle to create Forecast.",
      status: "completed",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 120,
      completedAtUnixMs: 120,
      errorMessage: null,
      steps: [
        {
          stepId: "plan-sheet-create",
          label: "Plan sheet creation",
          status: "completed",
          summary: "Prepared the semantic sheet-creation command for Forecast.",
          updatedAtUnixMs: 110,
        },
      ],
      artifact: {
        kind: "markdown",
        title: "Create Sheet Preview",
        text: "## Create Sheet Preview",
      },
    });

    expect(decoded.workflowTemplate).toBe("createSheet");
    expect(decoded.artifact?.title).toBe("Create Sheet Preview");
  });

  it("accepts cancelled workflow runs", () => {
    const decoded = decodeUnknownSync(WorkbookAgentWorkflowRunSchema, {
      runId: "workflow-4",
      threadId: "thr-1",
      startedByUserId: "alex@example.com",
      workflowTemplate: "summarizeWorkbook",
      title: "Summarize Workbook",
      summary: "Cancelled workflow: Summarize Workbook",
      status: "cancelled",
      createdAtUnixMs: 100,
      updatedAtUnixMs: 125,
      completedAtUnixMs: 125,
      errorMessage: "Cancelled by alex@example.com.",
      steps: [
        {
          stepId: "inspect-workbook",
          label: "Inspect workbook structure",
          status: "cancelled",
          summary: "Workflow cancelled before this step completed.",
          updatedAtUnixMs: 125,
        },
      ],
      artifact: null,
    });

    expect(decoded.status).toBe("cancelled");
    expect(decoded.steps[0]?.status).toBe("cancelled");
  });
});
