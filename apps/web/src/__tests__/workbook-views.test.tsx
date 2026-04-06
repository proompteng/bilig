// @vitest-environment jsdom
import { act } from "react";
import type { Viewport } from "@bilig/protocol";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  buildWorkbookSheetViewMutationArgs,
  useWorkbookViewsPane,
} from "../use-workbook-views-pane.js";
import type { WorkerRuntimeSelection } from "../runtime-session.js";
import type { ZeroWorkbookSheetViewSource } from "../use-workbook-views-pane.js";

interface MockZeroSheetViewHarness {
  readonly zero: ZeroWorkbookSheetViewSource;
  readonly mutations: unknown[];
  emit(value: unknown): void;
}

function createMockZeroSheetViewHarness(initialValue: unknown): MockZeroSheetViewHarness {
  let currentValue = initialValue;
  const listeners = new Set<(value: unknown) => void>();
  const mutations: unknown[] = [];
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
      mutate(mutation: unknown) {
        mutations.push(mutation);
        return {};
      },
    },
    mutations,
    emit(value: unknown) {
      currentValue = value;
      listeners.forEach((listener) => listener(value));
    },
  };
}

function ViewsHarness(props: {
  documentId: string;
  currentUserId: string;
  selection: WorkerRuntimeSelection;
  sheetNames: readonly string[];
  zero: ZeroWorkbookSheetViewSource;
  enabled: boolean;
  viewport: Viewport;
  onApply: (view: { sheetName: string | null; address: string; viewport: Viewport }) => void;
}) {
  const { viewsPanel, viewsToggle } = useWorkbookViewsPane({
    documentId: props.documentId,
    currentUserId: props.currentUserId,
    selection: props.selection,
    sheetNames: props.sheetNames,
    zero: props.zero,
    enabled: props.enabled,
    getCurrentViewport: () => props.viewport,
    onApply: props.onApply,
  });

  return (
    <div>
      {viewsToggle}
      {viewsPanel}
    </div>
  );
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

describe("workbook views", () => {
  it("saves the current workbook frame and applies accessible views", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const views = createMockZeroSheetViewHarness([
      {
        id: "view-shared",
        ownerUserId: "sam@example.com",
        name: "Shared QA",
        visibility: "shared",
        sheetId: 1,
        sheetName: "Sheet1",
        address: "B7",
        viewportJson: {
          rowStart: 6,
          rowEnd: 24,
          colStart: 1,
          colEnd: 8,
        },
        updatedAt: Date.parse("2026-04-06T13:00:00.000Z"),
      },
    ]);
    const onApply = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ViewsHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onApply={onApply}
          selection={{ sheetName: "Sheet1", address: "C4" }}
          sheetNames={["Sheet1"]}
          viewport={{ rowStart: 2, rowEnd: 18, colStart: 2, colEnd: 10 }}
          zero={views.zero}
        />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-views-toggle']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const input = host.querySelector<HTMLInputElement>("[data-testid='workbook-view-name']");
    expect(input).not.toBeNull();

    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set?.call(
        input,
        "Current focus",
      );
      input!.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-view-save']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(views.mutations).toHaveLength(1);
    expect(
      buildWorkbookSheetViewMutationArgs({
        documentId: "doc-1",
        id: "view-1",
        name: "Current focus",
        visibility: "private",
        selection: { sheetName: "Sheet1", address: "C4" },
        viewport: {
          rowStart: 2,
          rowEnd: 18,
          colStart: 2,
          colEnd: 10,
        },
      }),
    ).toEqual({
      documentId: "doc-1",
      id: "view-1",
      name: "Current focus",
      visibility: "private",
      sheetName: "Sheet1",
      address: "C4",
      viewport: {
        rowStart: 2,
        rowEnd: 18,
        colStart: 2,
        colEnd: 10,
      },
    });

    const applyButton = Array.from(host.querySelectorAll("button")).find(
      (button) => button.textContent === "Apply",
    );

    await act(async () => {
      applyButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onApply).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetName: "Sheet1",
        address: "B7",
        viewport: {
          rowStart: 6,
          rowEnd: 24,
          colStart: 1,
          colEnd: 8,
        },
      }),
    );

    await act(async () => {
      root.unmount();
    });
  });

  it("updates the visible view count when the Zero view publishes new rows", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const views = createMockZeroSheetViewHarness([]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ViewsHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          onApply={() => {}}
          selection={{ sheetName: "Sheet1", address: "A1" }}
          sheetNames={["Sheet1"]}
          viewport={{ rowStart: 0, rowEnd: 20, colStart: 0, colEnd: 10 }}
          zero={views.zero}
        />,
      );
    });

    expect(host.querySelector("[data-testid='workbook-views-toggle']")?.textContent).toContain("0");

    await act(async () => {
      views.emit([
        {
          id: "view-2",
          ownerUserId: "alex@example.com",
          name: "Daily review",
          visibility: "private",
          sheetId: 1,
          sheetName: "Sheet1",
          address: "F12",
          viewportJson: {
            rowStart: 8,
            rowEnd: 24,
            colStart: 5,
            colEnd: 12,
          },
          updatedAt: Date.now(),
        },
      ]);
    });

    expect(host.querySelector("[data-testid='workbook-views-toggle']")?.textContent).toContain("1");

    await act(async () => {
      root.unmount();
    });
  });
});
