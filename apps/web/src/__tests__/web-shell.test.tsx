// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { App } from "../App";

class ResizeObserverMock {
  observe() {}
  unobserve() {}
  disconnect() {}
}

function ensureCanvasAndResizeMocks() {
  Object.defineProperty(globalThis, "ResizeObserver", {
    configurable: true,
    value: ResizeObserverMock
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
                actualBoundingBoxDescent: 2
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
          }
        }
      )
  });
}

describe("web shell", () => {
  it("renders the minimal product shell without playground chrome", async () => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    ensureCanvasAndResizeMocks();

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<App />);
    });

    expect(host.querySelector("[data-testid='formula-bar']")).not.toBeNull();
    expect(host.querySelector("[data-testid='sheet-grid']")).not.toBeNull();
    expect(host.querySelector("[data-testid='preset-strip']")).toBeNull();
    expect(host.querySelector("[data-testid='metrics-panel']")).toBeNull();
    expect(host.querySelector("[data-testid='replica-panel']")).toBeNull();
    expect(host.textContent).not.toContain("Excel-scale shell on top of the local-first engine");

    await act(async () => {
      root.unmount();
    });
    host.remove();
  });
});
