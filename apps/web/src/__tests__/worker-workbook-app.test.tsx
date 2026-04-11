// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import type { Action, ToastT, ToastToDismiss } from "sonner";
import { toast } from "sonner";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkerWorkbookApp } from "../WorkerWorkbookApp.js";

async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve();
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

function findActiveToast(id: string): ToastT | null {
  return (
    toast
      .getToasts()
      .find(
        (entry: ToastT | ToastToDismiss): entry is ToastT =>
          !("dismiss" in entry) && entry.id === id,
      ) ?? null
  );
}

function getToastAction(activeToast: ToastT | null): Action {
  if (
    !activeToast ||
    !activeToast.action ||
    typeof activeToast.action !== "object" ||
    !("onClick" in activeToast.action)
  ) {
    throw new Error("Expected toast action object");
  }
  return activeToast.action;
}

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
    toast.dismiss();
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
      localPersistenceMode: "ephemeral",
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
    await flushToasts();

    const pendingToast = findActiveToast("pending-mutation-pending-1");
    expect(pendingToast?.title).toBe(
      "A local change could not be synced. mutation rejected by server",
    );
    const retryAction = getToastAction(pendingToast);

    await act(async () => {
      Reflect.apply(retryAction.onClick, undefined, [new MouseEvent("click")]);
    });

    expect(retryFailedPendingMutation).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("renders follower mode messaging when another tab owns persistent local storage", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      localPersistenceMode: "follower",
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

    expect(host.textContent).toContain("Another tab owns this workbook's persistent local store.");
    expect(host.textContent).toContain("following live document state");

    await act(async () => {
      root.unmount();
    });
  });
});
