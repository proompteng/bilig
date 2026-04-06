// @vitest-environment jsdom
import { act } from "react";
import type { Viewport } from "@bilig/protocol";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  useWorkbookScenariosPane,
  type ZeroWorkbookScenarioSource,
} from "../use-workbook-scenarios-pane.js";
import type { WorkerRuntimeSelection } from "../runtime-session.js";

interface MockZeroWorkbookScenarioHarness {
  readonly zero: ZeroWorkbookScenarioSource;
  emit(value: unknown): void;
}

function createMockZeroWorkbookScenarioHarness(
  initialValue: unknown,
): MockZeroWorkbookScenarioHarness {
  let currentValue = initialValue;
  const listeners = new Set<(value: unknown) => void>();
  const materializedView = {
    get data() {
      return currentValue;
    },
    addListener(listener: (value: unknown) => void) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    destroy() {},
  };

  return {
    zero: {
      materialize() {
        return materializedView;
      },
    },
    emit(value: unknown) {
      currentValue = value;
      listeners.forEach((listener) => listener(value));
    },
  };
}

function ScenariosHarness(props: {
  documentId: string;
  currentUserId: string;
  selection: WorkerRuntimeSelection;
  zero: ZeroWorkbookScenarioSource;
  enabled: boolean;
  viewport: Viewport;
  createScenario: (input: {
    documentId: string;
    name: string;
    selection: WorkerRuntimeSelection;
    viewport: Viewport;
  }) => Promise<{ documentId: string }>;
  deleteScenario: (input: { documentId: string; scenarioDocumentId: string }) => Promise<void>;
  navigateToWorkbook: (input: {
    documentId: string;
    sheetName?: string | null;
    address?: string | null;
  }) => void;
}) {
  const { scenariosPanel, scenariosToggle } = useWorkbookScenariosPane({
    documentId: props.documentId,
    currentUserId: props.currentUserId,
    selection: props.selection,
    zero: props.zero,
    enabled: props.enabled,
    getCurrentViewport: () => props.viewport,
    createScenario: props.createScenario,
    deleteScenario: props.deleteScenario,
    navigateToWorkbook: props.navigateToWorkbook,
  });

  return (
    <div>
      {scenariosToggle}
      {scenariosPanel}
    </div>
  );
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

describe("workbook scenarios", () => {
  it("creates a scratchpad branch from the current selection and navigates into it", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const scenarios = createMockZeroWorkbookScenarioHarness([]);
    const createScenario = vi.fn(async () => ({ documentId: "scenario:new" }));
    const deleteScenario = vi.fn(async () => undefined);
    const navigateToWorkbook = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ScenariosHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Revenue", address: "D12" }}
          viewport={{ rowStart: 4, rowEnd: 22, colStart: 2, colEnd: 10 }}
          zero={scenarios.zero}
          createScenario={createScenario}
          deleteScenario={deleteScenario}
          navigateToWorkbook={navigateToWorkbook}
        />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-scenarios-toggle']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const input = host.querySelector<HTMLInputElement>("[data-testid='workbook-scenario-name']");
    expect(input).not.toBeNull();

    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set?.call(
        input,
        "Margin downside",
      );
      input!.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-scenario-create']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(createScenario).toHaveBeenCalledWith({
      documentId: "doc-1",
      name: "Margin downside",
      selection: { sheetName: "Revenue", address: "D12" },
      viewport: {
        rowStart: 4,
        rowEnd: 22,
        colStart: 2,
        colEnd: 10,
      },
    });
    expect(navigateToWorkbook).toHaveBeenCalledWith({
      documentId: "scenario:new",
      sheetName: "Revenue",
      address: "D12",
    });

    await act(async () => {
      root.unmount();
    });
  });

  it("opens and deletes existing scratchpad branches from the live Zero list", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const scenarios = createMockZeroWorkbookScenarioHarness([
      {
        documentId: "scenario:existing",
        workbookId: "doc-1",
        ownerUserId: "alex@example.com",
        name: "Pricing what-if",
        baseRevision: 14,
        sheetId: 1,
        sheetName: "Pricing",
        address: "B7",
        viewportJson: {
          rowStart: 1,
          rowEnd: 18,
          colStart: 1,
          colEnd: 8,
        },
        createdAt: Date.parse("2026-04-06T10:00:00.000Z"),
        updatedAt: Date.parse("2026-04-06T10:00:00.000Z"),
      },
    ]);
    const createScenario = vi.fn(async () => ({ documentId: "scenario:new" }));
    const deleteScenario = vi.fn(async () => undefined);
    const navigateToWorkbook = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ScenariosHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Sheet1", address: "A1" }}
          viewport={{ rowStart: 0, rowEnd: 20, colStart: 0, colEnd: 10 }}
          zero={scenarios.zero}
          createScenario={createScenario}
          deleteScenario={deleteScenario}
          navigateToWorkbook={navigateToWorkbook}
        />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-scenarios-toggle']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const buttons = Array.from(host.querySelectorAll("button"));
    const openButton = buttons.find((button) => button.textContent === "Open");
    const deleteButton = buttons.find((button) => button.textContent === "Delete");

    await act(async () => {
      openButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    expect(navigateToWorkbook).toHaveBeenCalledWith({
      documentId: "scenario:existing",
      sheetName: "Pricing",
      address: "B7",
    });

    await act(async () => {
      deleteButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    expect(deleteScenario).toHaveBeenCalledWith({
      documentId: "doc-1",
      scenarioDocumentId: "scenario:existing",
    });

    await act(async () => {
      root.unmount();
    });
  });
});
