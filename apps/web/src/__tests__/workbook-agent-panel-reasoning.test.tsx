// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { WorkbookAgentPanel } from "../WorkbookAgentPanel.js";

beforeEach(() => {
  (
    globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
  ).IS_REACT_ACT_ENVIRONMENT = true;
});

afterEach(() => {
  document.body.innerHTML = "";
});

function renderPanel(entry: {
  id: string;
  kind: "plan" | "system";
  text: string | null;
}) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);

  return {
    host,
    root,
    render: async () => {
      await act(async () => {
        root.render(
          <WorkbookAgentPanel
            activeThreadId="thr-1"
            canFinalizeSharedBundle={false}
            canRecommendSharedBundle={false}
            cancellingWorkflowRunId={null}
            currentContext={{
              selection: { sheetName: "Sheet1", address: "A1" },
              viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
            }}
            currentUserSharedRecommendation={null}
            draft=""
            executionRecords={[]}
            isApplyingBundle={false}
            isLoading={false}
            isStartingWorkflow={false}
            onApplyPendingBundle={() => {}}
            onCancelWorkflowRun={() => {}}
            onDismissPendingBundle={() => {}}
            onDraftChange={() => {}}
            onInterrupt={() => {}}
            onReplayExecutionRecord={() => {}}
            onReviewPendingBundle={() => {}}
            onSelectAllPendingCommands={() => {}}
            onSelectThread={() => {}}
            onSelectThreadScope={() => {}}
            onStartNamedWorkflow={() => {}}
            onStartNewThread={() => {}}
            onStartSearchWorkflow={() => {}}
            onStartStructuralWorkflow={() => {}}
            onStartWorkflow={() => {}}
            onSubmit={() => {}}
            onTogglePendingCommand={() => {}}
            pendingBundle={null}
            preview={null}
            selectedCommandIndexes={[]}
            sharedApprovalOwnerUserId={null}
            sharedReviewDecidedByUserId={null}
            sharedReviewOwnerUserId={null}
            sharedReviewRecommendations={[]}
            sharedReviewStatus={null}
            snapshot={{
              sessionId: "agent-session-1",
              documentId: "doc-1",
              threadId: "thr-1",
              scope: "private",
              status: "idle",
              activeTurnId: null,
              lastError: null,
              context: {
                selection: { sheetName: "Sheet1", address: "A1" },
                viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
              },
              entries: [
                {
                  id: entry.id,
                  kind: entry.kind,
                  turnId: "turn-1",
                  text: entry.text,
                  phase: entry.kind === "plan" ? "reasoning" : null,
                  toolName: null,
                  toolStatus: null,
                  argumentsText: null,
                  outputText: null,
                  success: null,
                  citations: [],
                },
              ],
              pendingBundle: null,
              executionRecords: [],
              workflowRuns: [],
            }}
            threadScope="private"
            threadSummaries={[]}
            workflowRuns={[]}
          />,
        );
      });
    },
  };
}

describe("WorkbookAgentPanel reasoning", () => {
  it("renders plan entries as collapsible thought rows", async () => {
    const panel = renderPanel({
      id: "plan-1",
      kind: "plan",
      text: "Check the visible formulas before applying edits.",
    });

    await panel.render();

    expect(panel.host.textContent).toContain("Thought");
    expect(panel.host.textContent).toContain("Check the visible formulas before applying edits.");

    const toggle = panel.host.querySelector("[data-testid='workbook-agent-reasoning-toggle-plan-1']");
    expect(toggle instanceof HTMLButtonElement).toBe(true);

    await act(async () => {
      if (!(toggle instanceof HTMLButtonElement)) {
        throw new Error("Reasoning toggle not found");
      }
      toggle.click();
    });

    expect(
      panel.host.querySelector("[data-testid='workbook-agent-reasoning-panel-plan-1']"),
    ).not.toBeNull();

    await act(async () => {
      panel.root.unmount();
    });
  });

  it("hides legacy reasoning placeholders with no details", async () => {
    const panel = renderPanel({
      id: "system-1",
      kind: "system",
      text: "Codex emitted reasoning.",
    });

    await panel.render();

    expect(panel.host.textContent).not.toContain("Thought");
    expect(panel.host.textContent).not.toContain("Codex emitted reasoning.");

    await act(async () => {
      panel.root.unmount();
    });
  });
});
