// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it, vi } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { encodeViewportPatch } from "@bilig/worker-transport";
import { App } from "../App";

const fakeClient = {
  async invoke(method: string) {
    switch (method) {
      case "bootstrap":
      case "getRuntimeState":
        return {
          workbookName: "bilig-demo",
          sheetNames: ["Sheet1"],
          metrics: {
            batchId: 0,
            changedInputCount: 0,
            dirtyFormulaCount: 0,
            wasmFormulaCount: 0,
            jsFormulaCount: 0,
            rangeNodeVisits: 0,
            recalcMs: 0,
            compileMs: 0,
          },
          syncState: "local-only",
        };
      case "getCell":
        return {
          sheetName: "Sheet1",
          address: "A1",
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 0,
        };
      case "updateColumnWidth":
      case "autofitColumn":
        return 104;
      default:
        return undefined;
    }
  },
  ready: async () => undefined,
  subscribe: () => () => {},
  subscribeBatches: () => () => {},
  subscribeViewportPatches: (_subscription: unknown, listener: (bytes: Uint8Array) => void) => {
    listener(
      encodeViewportPatch({
        version: 1,
        full: true,
        viewport: {
          sheetName: "Sheet1",
          rowStart: 0,
          rowEnd: 23,
          colStart: 0,
          colEnd: 11,
        },
        metrics: {
          batchId: 0,
          changedInputCount: 0,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
        styles: [],
        cells: [],
        columns: [],
        rows: [],
      }),
    );
    return () => {};
  },
  dispose() {},
};

vi.mock("@bilig/worker-transport", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@bilig/worker-transport")>();
  return {
    ...actual,
    createWorkerEngineClient: () => fakeClient,
  };
});

vi.mock("@rocicorp/zero/react", () => ({
  useZero: () => ({
    mutate: () => ({
      client: Promise.resolve({ type: "complete" }),
    }),
    materialize: () => ({
      data: undefined,
      addListener: () => () => {},
      destroy() {},
    }),
  }),
  useConnectionState: () => ({ name: "connected" }),
}));

class ResizeObserverMock {
  observe() {}
  unobserve() {}
  disconnect() {}
}

function ensureCanvasAndResizeMocks() {
  Object.defineProperty(globalThis, "ResizeObserver", {
    configurable: true,
    value: ResizeObserverMock,
  });

  Object.defineProperty(HTMLCanvasElement.prototype, "getContext", {
    configurable: true,
    value: () =>
      new Proxy(
        {},
        {
          get(_target, property) {
            if (property === "measureText") {
              return (value: string) => ({
                width: value.length * 8,
                actualBoundingBoxAscent: 8,
                actualBoundingBoxDescent: 2,
              });
            }
            if (property === "createLinearGradient" || property === "createPattern") {
              return () => ({ addColorStop() {} });
            }
            if (property === "getImageData") {
              return () => ({ data: new Uint8ClampedArray(4) });
            }
            return () => {};
          },
          set() {
            return true;
          },
        },
      ),
  });
}

function waitFor(host: HTMLElement, predicate: () => boolean, attempts = 30): Promise<void> {
  const poll = async (remaining: number): Promise<void> => {
    if (predicate()) {
      return;
    }
    if (remaining <= 0) {
      const errorText = host.querySelector("[data-testid='worker-error']")?.textContent;
      throw new Error(errorText ?? host.innerHTML ?? "Timed out waiting for worker-backed shell");
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
    await poll(remaining - 1);
  };
  return poll(attempts);
}

class WorkerMock {
  constructor(...args: unknown[]) {
    void args;
  }

  postMessage(): void {}

  addEventListener(): void {}

  removeEventListener(): void {}

  terminate(): void {}
}

describe("web shell", () => {
  it("renders the minimal product shell without playground chrome", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;
    ensureCanvasAndResizeMocks();
    Object.defineProperty(globalThis, "Worker", {
      configurable: true,
      value: WorkerMock,
    });

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <App
          config={{
            apiBaseUrl: "http://127.0.0.1:4321",
            zeroCacheUrl: "http://127.0.0.1:4848",
            defaultDocumentId: "bilig-demo",
            persistState: true,
            zeroViewportBridge: false,
          }}
        />,
      );
    });

    await act(async () => {
      await waitFor(host, () => host.querySelector("[data-testid='formula-bar']") !== null);
    });

    expect(host.querySelector("[data-testid='formula-bar']")).not.toBeNull();
    expect(host.querySelector("[data-testid='sheet-grid']")).not.toBeNull();
    expect(host.querySelector("[data-testid='preset-strip']")).toBeNull();
    expect(host.querySelector("[data-testid='metrics-panel']")).toBeNull();
    expect(host.querySelector("[data-testid='replica-panel']")).toBeNull();
    expect(host.querySelector("h1")).toBeNull();
    expect(host.textContent).not.toContain("Excel-scale shell on top of the local-first engine");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
