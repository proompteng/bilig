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
      artifact: {
        kind: "markdown",
        title: "Workbook Summary",
        text: "## Summary",
      },
    });

    expect(decoded.workflowTemplate).toBe("summarizeWorkbook");
    expect(decoded.artifact).toEqual(
      expect.objectContaining({
        kind: "markdown",
        title: "Workbook Summary",
      }),
    );
  });
});
