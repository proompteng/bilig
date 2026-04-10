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
          summary: 'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 2 matches.',
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
});
