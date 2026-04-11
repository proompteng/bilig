// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkbookAgentPanel } from "../WorkbookAgentPanel.js";

function createSnapshot(overrides: Record<string, unknown> = {}) {
  return {
    sessionId: "agent-session-1",
    documentId: "doc-1",
    threadId: "thr-1",
    scope: "private",
    status: "idle",
    activeTurnId: null,
    lastError: null,
    context: {
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
    entries: [],
    pendingBundle: null,
    executionRecords: [],
    workflowRuns: [],
    ...overrides,
  };
}

function renderPanel(overrides: Record<string, unknown> = {}) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  const props = {
    activeThreadId: "thr-1",
    currentContext: null,
    snapshot: createSnapshot(),
    pendingBundle: null,
    preview: null,
    sharedApprovalOwnerUserId: null,
    sharedReviewOwnerUserId: null,
    sharedReviewStatus: null,
    sharedReviewDecidedByUserId: null,
    sharedReviewRecommendations: [],
    currentUserSharedRecommendation: null,
    canFinalizeSharedBundle: false,
    canRecommendSharedBundle: false,
    selectedCommandIndexes: [],
    executionRecords: [],
    workflowRuns: [],
    cancellingWorkflowRunId: null,
    isStartingWorkflow: false,
    threadScope: "private",
    threadSummaries: [],
    draft: "",
    isLoading: false,
    isApplyingBundle: false,
    onApplyPendingBundle: vi.fn(),
    onDraftChange: vi.fn(),
    onDismissPendingBundle: vi.fn(),
    onReviewPendingBundle: vi.fn(),
    onInterrupt: vi.fn(),
    onSelectAllPendingCommands: vi.fn(),
    onSelectThreadScope: vi.fn(),
    onSelectThread: vi.fn(),
    onStartNewThread: vi.fn(),
    onTogglePendingCommand: vi.fn(),
    onCancelWorkflowRun: vi.fn(),
    onReplayExecutionRecord: vi.fn(),
    onStartWorkflow: vi.fn(),
    onStartNamedWorkflow: vi.fn(),
    onStartSearchWorkflow: vi.fn(),
    onStartStructuralWorkflow: vi.fn(),
    onSubmit: vi.fn(),
    ...overrides,
  };

  return {
    host,
    root,
    render: async () => {
      await act(async () => {
        root.render(<WorkbookAgentPanel {...props} />);
      });
    },
    unmount: async () => {
      await act(async () => {
        root.unmount();
      });
    },
  };
}

afterEach(() => {
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

describe("workbook agent markdown rendering", () => {
  it("renders assistant markdown emphasis instead of raw markers", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const snapshot = createSnapshot({
      entries: [
        {
          id: "assistant-1",
          kind: "assistant",
          turnId: "turn-1",
          text: "I mean the **staged changes / preview side rail** in the workbook UI, with **Apply** or **Dismiss**.",
          phase: null,
          toolName: null,
          toolStatus: null,
          argumentsText: null,
          outputText: null,
          success: null,
          citations: [],
        },
      ],
    });
    const panel = renderPanel({ snapshot });

    await panel.render();

    expect(panel.host.textContent).toContain("staged changes / preview side rail");
    expect(panel.host.textContent).not.toContain("**staged changes / preview side rail**");
    const strongNodes = panel.host.querySelectorAll("strong");
    expect(strongNodes.length).toBeGreaterThanOrEqual(2);
    expect([...strongNodes].some((node) => node.textContent === "Apply")).toBe(true);

    await panel.unmount();
  });

  it("renders workflow markdown artifacts as structured content without duplicating the title heading", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const panel = renderPanel({
      snapshot: createSnapshot({
        entries: [
          {
            id: "assistant-0",
            kind: "assistant",
            turnId: "turn-0",
            text: "Drafted a workbook summary.",
            phase: null,
            toolName: null,
            toolStatus: null,
            argumentsText: null,
            outputText: null,
            success: null,
            citations: [],
          },
        ],
      }),
      workflowRuns: [
        {
          runId: "wf-1",
          threadId: "thr-1",
          startedByUserId: "alex@example.com",
          workflowTemplate: "summarizeWorkbook",
          title: "Summarize Workbook",
          summary: "Summarized workbook structure across 2 sheets.",
          status: "completed",
          createdAtUnixMs: 1,
          updatedAtUnixMs: 2,
          completedAtUnixMs: 2,
          errorMessage: null,
          steps: [],
          artifact: {
            kind: "markdown",
            title: "Workbook Summary",
            text: "## Workbook Summary\n\nSheets: 2\n### Sheets\n- Sheet1\n- Sheet2",
          },
        },
      ],
    });

    await panel.render();

    expect(panel.host.textContent).toContain("Workbook Summary");
    expect(panel.host.textContent).toContain("Sheets: 2");
    expect(panel.host.textContent).toContain("Sheet1");
    expect(panel.host.textContent?.split("Workbook Summary").length).toBe(2);
    const listItems = [...panel.host.querySelectorAll("li")].map((node) => node.textContent);
    expect(listItems).toEqual(expect.arrayContaining(["Sheet1", "Sheet2"]));

    await panel.unmount();
  });
});
