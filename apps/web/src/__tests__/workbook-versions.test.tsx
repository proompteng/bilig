// @vitest-environment jsdom
import { act } from "react";
import type { Viewport } from "@bilig/protocol";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  buildWorkbookVersionMutationArgs,
  useWorkbookVersionsPane,
  type ZeroWorkbookVersionSource,
} from "../use-workbook-versions-pane.js";
import type { WorkerRuntimeSelection } from "../runtime-session.js";

interface MockZeroWorkbookVersionHarness {
  readonly zero: ZeroWorkbookVersionSource;
  readonly mutations: unknown[];
  emit(value: unknown): void;
}

function createMockZeroWorkbookVersionHarness(
  initialValue: unknown,
): MockZeroWorkbookVersionHarness {
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

function VersionsHarness(props: {
  documentId: string;
  currentUserId: string;
  selection: WorkerRuntimeSelection;
  zero: ZeroWorkbookVersionSource;
  enabled: boolean;
  viewport: Viewport;
}) {
  const { versionsPanel, versionsToggle } = useWorkbookVersionsPane({
    documentId: props.documentId,
    currentUserId: props.currentUserId,
    selection: props.selection,
    zero: props.zero,
    enabled: props.enabled,
    getCurrentViewport: () => props.viewport,
  });

  return (
    <div>
      {versionsToggle}
      {versionsPanel}
    </div>
  );
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

describe("workbook versions", () => {
  it("saves and restores authoritative workbook versions from the mounted panel", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const versions = createMockZeroWorkbookVersionHarness([
      {
        id: "version-restore",
        ownerUserId: "sam@example.com",
        name: "Month close",
        revision: 12,
        sheetId: 1,
        sheetName: "Sheet1",
        address: "D5",
        viewportJson: {
          rowStart: 4,
          rowEnd: 20,
          colStart: 3,
          colEnd: 12,
        },
        createdAt: Date.parse("2026-04-06T13:00:00.000Z"),
        updatedAt: Date.parse("2026-04-06T13:00:00.000Z"),
      },
    ]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <VersionsHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Sheet1", address: "B2" }}
          viewport={{ rowStart: 1, rowEnd: 14, colStart: 1, colEnd: 8 }}
          zero={versions.zero}
        />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-versions-toggle']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const input = host.querySelector<HTMLInputElement>("[data-testid='workbook-version-name']");
    expect(input).not.toBeNull();

    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set?.call(
        input,
        "Safe restore point",
      );
      input!.dispatchEvent(new Event("input", { bubbles: true }));
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-version-save']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(versions.mutations).toHaveLength(1);
    expect(
      buildWorkbookVersionMutationArgs({
        documentId: "doc-1",
        id: "version-1",
        name: "Safe restore point",
        selection: { sheetName: "Sheet1", address: "B2" },
        viewport: {
          rowStart: 1,
          rowEnd: 14,
          colStart: 1,
          colEnd: 8,
        },
      }),
    ).toEqual({
      documentId: "doc-1",
      id: "version-1",
      name: "Safe restore point",
      sheetName: "Sheet1",
      address: "B2",
      viewport: {
        rowStart: 1,
        rowEnd: 14,
        colStart: 1,
        colEnd: 8,
      },
    });

    const restoreButton = Array.from(host.querySelectorAll("button")).find(
      (button) => button.textContent === "Restore",
    );
    await act(async () => {
      restoreButton?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(versions.mutations).toHaveLength(2);

    await act(async () => {
      root.unmount();
    });
  });

  it("updates the visible version count when the Zero view publishes new rows", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const versions = createMockZeroWorkbookVersionHarness([]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <VersionsHarness
          currentUserId="alex@example.com"
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Sheet1", address: "A1" }}
          viewport={{ rowStart: 0, rowEnd: 20, colStart: 0, colEnd: 10 }}
          zero={versions.zero}
        />,
      );
    });

    expect(host.querySelector("[data-testid='workbook-versions-toggle']")?.textContent).toContain(
      "0",
    );

    await act(async () => {
      versions.emit([
        {
          id: "version-2",
          ownerUserId: "alex@example.com",
          name: "Review baseline",
          revision: 8,
          sheetId: 1,
          sheetName: "Sheet1",
          address: "F12",
          viewportJson: {
            rowStart: 8,
            rowEnd: 24,
            colStart: 5,
            colEnd: 12,
          },
          createdAt: Date.now(),
          updatedAt: Date.now(),
        },
      ]);
    });

    expect(host.querySelector("[data-testid='workbook-versions-toggle']")?.textContent).toContain(
      "1",
    );

    await act(async () => {
      root.unmount();
    });
  });
});
