// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkbookPresenceBar } from "../WorkbookPresenceBar.js";
import { useWorkbookPresence } from "../use-workbook-presence.js";
import {
  WORKBOOK_PRESENCE_HEARTBEAT_MS,
  WORKBOOK_PRESENCE_STALE_AFTER_MS,
} from "../workbook-presence-model.js";

interface MockZeroPresenceHarness {
  readonly mutateCalls: unknown[];
  readonly zero: {
    materialize(): {
      readonly data: unknown;
      addListener(listener: (value: unknown) => void): () => void;
      destroy(): void;
    };
    mutate(request: unknown): { client: Promise<{ type: "complete" }> };
  };
  emit(value: unknown): void;
}

function createMockZeroPresenceHarness(initialValue: unknown): MockZeroPresenceHarness {
  let currentValue = initialValue;
  const listeners = new Set<(value: unknown) => void>();
  const mutateCalls: unknown[] = [];
  const view = {
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
    mutateCalls,
    zero: {
      materialize() {
        return view;
      },
      mutate(request: unknown) {
        mutateCalls.push(request);
        return {
          client: Promise.resolve({ type: "complete" as const }),
        };
      },
    },
    emit(value: unknown) {
      currentValue = value;
      listeners.forEach((listener) => listener(value));
    },
  };
}

function PresenceHarness(props: {
  documentId: string;
  sessionId: string;
  selection: { sheetName: string; address: string };
  sheetNames: readonly string[];
  zero: MockZeroPresenceHarness["zero"];
  enabled: boolean;
  onJump: (sheetName: string, address: string) => void;
}) {
  const collaborators = useWorkbookPresence({
    documentId: props.documentId,
    sessionId: props.sessionId,
    selection: props.selection,
    sheetNames: props.sheetNames,
    zero: props.zero,
    enabled: props.enabled,
  });
  return <WorkbookPresenceBar collaborators={collaborators} onJump={props.onJump} />;
}

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  document.body.innerHTML = "";
});

describe("workbook presence", () => {
  it("publishes local selection presence and renders collaborator jump chips", async () => {
    const presence = createMockZeroPresenceHarness([
      {
        sessionId: "doc-1:browser:other",
        userId: "amy.smith@example.com",
        sheetId: 1,
        sheetName: "Sheet1",
        address: "B7",
        selectionJson: { sheetName: "Sheet1", address: "B7" },
        updatedAt: Date.now(),
      },
      {
        sessionId: "doc-1:browser:self",
        userId: "me@example.com",
        sheetId: 1,
        sheetName: "Sheet1",
        address: "A1",
        selectionJson: { sheetName: "Sheet1", address: "A1" },
        updatedAt: Date.now(),
      },
    ]);
    const onJump = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <PresenceHarness
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Sheet1", address: "A1" }}
          sessionId="doc-1:browser:self"
          sheetNames={["Sheet1"]}
          zero={presence.zero}
          onJump={onJump}
        />,
      );
    });

    expect(presence.mutateCalls).toHaveLength(1);
    expect(presence.mutateCalls[0]).toMatchObject({
      args: {
        documentId: "doc-1",
        sessionId: "doc-1:browser:self",
        sheetName: "Sheet1",
        address: "A1",
        selection: {
          sheetName: "Sheet1",
          address: "A1",
        },
      },
    });

    const chips = host.querySelectorAll("[data-testid='ax-presence-chip']");
    expect(chips).toHaveLength(1);
    expect(chips[0]?.textContent).toContain("Amy Smith");

    await act(async () => {
      chips[0]?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onJump).toHaveBeenCalledWith("Sheet1", "B7");

    await act(async () => {
      root.unmount();
    });
  });

  it("drops stale collaborators and keeps publishing heartbeat updates", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-06T12:00:00.000Z"));

    const presence = createMockZeroPresenceHarness([
      {
        sessionId: "doc-1:browser:other",
        userId: "guest:deadbeef",
        sheetId: 1,
        sheetName: "Sheet1",
        address: "C9",
        selectionJson: { sheetName: "Sheet1", address: "C9" },
        updatedAt: Date.now() - WORKBOOK_PRESENCE_STALE_AFTER_MS - 1,
      },
    ]);
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <PresenceHarness
          documentId="doc-1"
          enabled
          selection={{ sheetName: "Sheet1", address: "A1" }}
          sessionId="doc-1:browser:self"
          sheetNames={["Sheet1"]}
          zero={presence.zero}
          onJump={() => {}}
        />,
      );
    });

    expect(host.querySelector("[data-testid='ax-presence-chip']")).toBeNull();
    expect(presence.mutateCalls).toHaveLength(1);

    await act(async () => {
      vi.advanceTimersByTime(WORKBOOK_PRESENCE_HEARTBEAT_MS);
      await Promise.resolve();
    });

    expect(presence.mutateCalls).toHaveLength(2);

    await act(async () => {
      presence.emit([
        {
          sessionId: "doc-1:browser:other",
          userId: "guest:deadbeef",
          sheetId: 1,
          sheetName: "Missing",
          address: "C9",
          selectionJson: { sheetName: "Missing", address: "C9" },
          updatedAt: Date.now(),
        },
        {
          sessionId: "doc-1:browser:other-2",
          userId: "guest:facefeed",
          sheetId: 1,
          sheetName: "Sheet1",
          address: "D4",
          selectionJson: { sheetName: "Sheet1", address: "D4" },
          updatedAt: Date.now(),
        },
      ]);
    });

    const chips = host.querySelectorAll("[data-testid='ax-presence-chip']");
    expect(chips).toHaveLength(1);
    expect(chips[0]?.textContent).toContain("Guest FEED");

    await act(async () => {
      root.unmount();
    });
  });
});
