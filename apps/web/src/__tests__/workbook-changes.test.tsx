// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useWorkbookChangesPane } from "../use-workbook-changes-pane.js";
import type { ZeroWorkbookChangeSource } from "../use-workbook-changes-pane.js";

interface MockZeroChangeHarness {
  readonly zero: ZeroWorkbookChangeSource;
  readonly mutations: unknown[];
  readonly materializedView: {
    readonly data: unknown;
    addListener(listener: (value: unknown) => void): () => void;
    destroy(): void;
  };
  emit(value: unknown): void;
}

function createMockZeroChangeHarness(initialValue: unknown): MockZeroChangeHarness {
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
    materializedView,
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

function ChangesHarness(props: {
  documentId: string;
  sheetNames: readonly string[];
  zero: MockZeroChangeHarness["zero"];
  enabled: boolean;
  onJump: (sheetName: string, address: string) => void;
}) {
  const { changeCount, changesPanel } = useWorkbookChangesPane({
    documentId: props.documentId,
    sheetNames: props.sheetNames,
    zero: props.zero,
    enabled: props.enabled,
    onJump: props.onJump,
  });

  return (
    <div>
      <div data-testid="workbook-changes-count">{String(changeCount)}</div>
      {changesPanel}
    </div>
  );
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

describe("workbook changes", () => {
  it("renders authoritative change rows and jumps to available anchors", async () => {
    const changes = createMockZeroChangeHarness([
      {
        revision: 12,
        actorUserId: "amy.smith@example.com",
        clientMutationId: "mutation-12",
        eventKind: "fillRange",
        summary: "Filled Sheet1!B1:B3",
        sheetId: 1,
        sheetName: "Sheet1",
        anchorAddress: "B1",
        rangeJson: { sheetName: "Sheet1", startAddress: "B1", endAddress: "B3" },
        undoBundleJson: {
          kind: "engineOps",
          ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "B1" }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse("2026-04-06T12:34:00.000Z"),
      },
      {
        revision: 11,
        actorUserId: "guest:deadbeef",
        clientMutationId: null,
        eventKind: "renderCommit",
        summary: "Deleted sheet Archive",
        sheetId: null,
        sheetName: null,
        anchorAddress: null,
        rangeJson: null,
        undoBundleJson: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse("2026-04-06T12:30:00.000Z"),
      },
    ]);
    const onJump = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ChangesHarness
          documentId="doc-1"
          enabled
          onJump={onJump}
          sheetNames={["Sheet1"]}
          zero={changes.zero}
        />,
      );
    });

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe("2");

    const rows = host.querySelectorAll<HTMLElement>("[data-testid='workbook-change-row']");
    expect(rows).toHaveLength(2);
    expect(rows[0]?.textContent).toContain("Filled Sheet1!B1:B3");
    expect(rows[0]?.textContent).toContain("Amy Smith");
    expect(rows[1]?.textContent).toContain("Deleted sheet Archive");

    await act(async () => {
      rows[0]?.querySelector("button")?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onJump).toHaveBeenCalledWith("Sheet1", "B1");

    await act(async () => {
      root.unmount();
    });
  });

  it("updates the visible change count when the Zero view publishes new rows", async () => {
    const changes = createMockZeroChangeHarness([]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ChangesHarness
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={["Sheet1"]}
          zero={changes.zero}
        />,
      );
    });

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe("0");

    await act(async () => {
      changes.emit([
        {
          revision: 15,
          actorUserId: "alex@example.com",
          clientMutationId: "mutation-15",
          eventKind: "setCellValue",
          summary: "Updated Sheet1!C7",
          sheetId: 1,
          sheetName: "Sheet1",
          anchorAddress: "C7",
          rangeJson: { sheetName: "Sheet1", startAddress: "C7", endAddress: "C7" },
          undoBundleJson: {
            kind: "engineOps",
            ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "C7" }],
          },
          revertedByRevision: null,
          revertsRevision: null,
          createdAt: Date.now(),
        },
      ]);
    });

    expect(host.querySelector("[data-testid='workbook-changes-count']")?.textContent).toBe("1");

    await act(async () => {
      root.unmount();
    });
  });

  it("routes revert actions through the authoritative workbook change mutator", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    const changes = createMockZeroChangeHarness([
      {
        revision: 21,
        actorUserId: "alex@example.com",
        clientMutationId: "mutation-21",
        eventKind: "setCellValue",
        summary: "Updated Sheet1!A1",
        sheetId: 1,
        sheetName: "Sheet1",
        anchorAddress: "A1",
        rangeJson: { sheetName: "Sheet1", startAddress: "A1", endAddress: "A1" },
        undoBundleJson: {
          kind: "engineOps",
          ops: [{ kind: "clearCell", sheetName: "Sheet1", address: "A1" }],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAt: Date.parse("2026-04-06T13:12:00.000Z"),
      },
    ]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <ChangesHarness
          documentId="doc-1"
          enabled
          onJump={() => {}}
          sheetNames={["Sheet1"]}
          zero={changes.zero}
        />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='workbook-change-revert']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(changes.mutations).toHaveLength(1);

    await act(async () => {
      root.unmount();
    });
  });
});
