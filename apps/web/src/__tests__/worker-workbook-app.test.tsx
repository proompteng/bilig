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

const { latestWorkbookViewProps, useWorkerWorkbookAppState } = vi.hoisted(() => ({
  latestWorkbookViewProps: { current: null as Record<string, unknown> | null },
  useWorkerWorkbookAppState: vi.fn(),
}));

vi.mock("@bilig/grid", () => ({
  WorkbookView: (props: Record<string, unknown>) => {
    latestWorkbookViewProps.current = props;
    return <div data-testid="workbook-view" />;
  },
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
      approvePersistenceTransfer: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      pendingTransferRequest: null,
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation,
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: false,
      localPersistenceMode: "ephemeral",
      statusModeLabel: "Live",
      transferRequested: false,
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

  it("keeps follower banner hidden while live multi-tab sync is available", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const requestPersistenceTransfer = vi.fn();
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      approvePersistenceTransfer: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      pendingTransferRequest: null,
      requestPersistenceTransfer,
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      localPersistenceMode: "follower",
      statusModeLabel: "Live",
      transferRequested: false,
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

    expect(host.textContent).not.toContain("Another tab is the local writer.");
    expect(host.querySelector("button")?.textContent).not.toBe("Become writer");

    await act(async () => {
      root.unmount();
    });
  });

  it("renders follower controls when local persistence is degraded", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const requestPersistenceTransfer = vi.fn();
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      approvePersistenceTransfer: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      pendingTransferRequest: null,
      requestPersistenceTransfer,
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: false,
      runtimeError: null,
      runtimeReady: true,
      localPersistenceMode: "follower",
      statusModeLabel: "Live",
      transferRequested: false,
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

    expect(host.textContent).toContain("Another tab is the local writer.");
    expect(host.textContent).toContain("Become writer");

    await act(async () => {
      host.querySelector("button")?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(requestPersistenceTransfer).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("renders writer transfer controls when another tab requests persistence handoff", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const approvePersistenceTransfer = vi.fn();
    const dismissPersistenceTransferRequest = vi.fn();
    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer,
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      dismissPersistenceTransferRequest,
      editorConflictBanner: null,
      failedPendingMutation: null,
      pendingTransferRequest: {
        requestId: "req-1",
        requesterTabId: "tab:other",
        requestedAtUnixMs: Date.now(),
      },
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      localPersistenceMode: "persistent",
      statusModeLabel: "Live",
      transferRequested: false,
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

    expect(host.textContent).toContain("Another tab wants writer ownership for local storage.");
    const buttons = [...host.querySelectorAll("button")];
    expect(buttons.map((button) => button.textContent)).toEqual(["Transfer writer", "Stay writer"]);

    await act(async () => {
      buttons[0]?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      buttons[1]?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(approvePersistenceTransfer).toHaveBeenCalledTimes(1);
    expect(dismissPersistenceTransferRequest).toHaveBeenCalledTimes(1);

    await act(async () => {
      root.unmount();
    });
  });

  it("passes an authoritative selection-range callback into the workbook view", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    useWorkerWorkbookAppState.mockReturnValue({
      agentError: null,
      approvePersistenceTransfer: vi.fn(),
      clearAgentError: vi.fn(),
      clearRuntimeError: vi.fn(),
      dismissPersistenceTransferRequest: vi.fn(),
      editorConflictBanner: null,
      failedPendingMutation: null,
      pendingTransferRequest: null,
      requestPersistenceTransfer: vi.fn(),
      retryFailedPendingMutation: vi.fn(),
      remoteSyncAvailable: true,
      runtimeError: null,
      runtimeReady: true,
      workbookReady: true,
      zeroConfigured: true,
      localPersistenceMode: "persistent",
      statusModeLabel: "Live",
      transferRequested: false,
      workerHandle: {
        viewportStore: {},
      },
      handleSelectionRangeChange: vi.fn(),
      selection: { sheetName: "Sheet1", address: "B18" },
      selectedCell: { sheetName: "Sheet1", address: "B18" },
      sheetNames: ["Sheet1"],
      previewRanges: [],
      resolvedValue: "",
      sidePanel: null,
      sidePanelId: undefined,
      sidePanelWidth: undefined,
      setSidePanelWidth: vi.fn(),
      commitEditor: vi.fn(),
      copySelectionRange: vi.fn(),
      createSheet: vi.fn(),
      fillSelectionRange: vi.fn(),
      handleEditorChange: vi.fn(),
      handleVisibleViewportChange: vi.fn(),
      invokeColumnVisibilityMutation: vi.fn(),
      invokeDeleteColumnsMutation: vi.fn(),
      invokeDeleteRowsMutation: vi.fn(),
      invokeInsertColumnsMutation: vi.fn(),
      invokeInsertRowsMutation: vi.fn(),
      invokeRowVisibilityMutation: vi.fn(),
      invokeSetFreezePaneMutation: vi.fn(),
      isEditing: false,
      isEditingCell: false,
      moveSelectionRange: vi.fn(),
      pasteIntoSelection: vi.fn(),
      renameSheet: vi.fn(),
      reportRuntimeError: vi.fn(),
      runtimeErrorBanner: null,
      selectAddress: vi.fn(),
      approveBundle: vi.fn(),
      ribbon: null,
      dismissPersistenceTransferRequestBanner: null,
      subscribeViewport: vi.fn(),
      columnWidths: {},
      hiddenColumns: {},
      hiddenRows: {},
      rowHeights: {},
      freezeRows: 0,
      freezeCols: 0,
      toggleBooleanCell: vi.fn(),
      visibleEditorValue: "",
      importPanel: null,
      importToggle: null,
      clearImportError: vi.fn(),
      pendingCommandCount: 0,
      canUndo: false,
      canRedo: false,
      changeCount: 0,
      changesPanel: null,
      redoLatestChange: vi.fn(),
      undoLatestChange: vi.fn(),
      startNewThread: vi.fn(),
      requestPersistenceTransferBanner: null,
      localPersistenceBanner: null,
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

    expect(typeof latestWorkbookViewProps.current?.["onSelectionRangeChange"]).toBe("function");

    await act(async () => {
      root.unmount();
    });
  });
});
