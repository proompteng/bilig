// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { useWorkbookChangesPane } from "../use-workbook-changes-pane.js";
import type { ZeroWorkbookChangeQuerySource } from "../use-workbook-changes.js";

interface MockZeroChangeHarness {
  readonly zero: ZeroWorkbookChangeQuerySource;
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
    },
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
  const { changesPanel, changesToggle } = useWorkbookChangesPane({
    documentId: props.documentId,
    sheetNames: props.sheetNames,
    zero: props.zero,
    enabled: props.enabled,
    onJump: props.onJump,
  });

  return (
    <div>
      {changesToggle}
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

    const toggle = host.querySelector("[data-testid='workbook-changes-toggle']");
    expect(toggle?.textContent).toContain("2");

    await act(async () => {
      toggle?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const rows = host.querySelectorAll<HTMLButtonElement>("[data-testid='workbook-change-row']");
    expect(rows).toHaveLength(2);
    expect(rows[0]?.textContent).toContain("Filled Sheet1!B1:B3");
    expect(rows[0]?.textContent).toContain("Amy Smith");
    expect(rows[1]?.textContent).toContain("Deleted sheet Archive");
    expect(rows[1]?.disabled).toBe(true);

    await act(async () => {
      rows[0]?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
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

    expect(host.querySelector("[data-testid='workbook-changes-toggle']")?.textContent).toContain(
      "0",
    );

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
          createdAt: Date.now(),
        },
      ]);
    });

    expect(host.querySelector("[data-testid='workbook-changes-toggle']")?.textContent).toContain(
      "1",
    );

    await act(async () => {
      root.unmount();
    });
  });
});
