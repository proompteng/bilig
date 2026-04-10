// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkerWorkbookApp } from "../WorkerWorkbookApp.js";

const { useWorkerWorkbookAppState } = vi.hoisted(() => ({
  useWorkerWorkbookAppState: vi.fn(),
}));

vi.mock("@bilig/grid", () => ({
  WorkbookView: () => <div data-testid="workbook-view" />,
}));

vi.mock("../use-worker-workbook-app-state.js", () => ({
  useWorkerWorkbookAppState,
}));

vi.mock("../use-workbook-import-pane.js", () => ({
  useWorkbookImportPane: () => ({
    clearImportError: vi.fn(),
    importError: null,
    importPanel: null,
    importToggle: null,
  }),
}));

vi.mock("../use-workbook-shortcut-dialog.js", () => ({
  useWorkbookShortcutDialog: () => ({
    shortcutDialog: null,
    shortcutHelpButton: null,
  }),
}));

describe("WorkerWorkbookApp", () => {
  afterEach(() => {
    vi.clearAllMocks();
    document.body.innerHTML = "";
  });

  it("renders a retry toast for failed pending mutations", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const retryFailedPendingMutation = vi.fn();
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: {
        id: "pending-1",
        method: "setCellValue",
        failureMessage: "mutation rejected by server",
        attemptCount: 1,
      },
      retryFailedPendingMutation,
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: false,
      statusModeLabel: "Live",
      workbookReady: false,
      workerHandle: null,
      zeroConfigured: true,
    });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkerWorkbookApp
          config={{
            currentUserId: "guest:test",
            defaultDocumentId: "doc-1",
            persistState: true,
            zeroCacheUrl: "http://127.0.0.1:4848",
          }}
          connectionState={{ name: "connected" }}
        />,
      );
    });

    expect(host.textContent).toContain("A local change could not be synced.");
    expect(host.textContent).toContain("mutation rejected by server");

    const retryButton = host.querySelector(
      "[data-testid='workbook-toast-pending-mutation-pending-1-action']",
    );
    expect(retryButton).not.toBeNull();

    await act(async () => {
      retryButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(retryFailedPendingMutation).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });
});
