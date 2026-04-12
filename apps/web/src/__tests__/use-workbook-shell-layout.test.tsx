// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH,
  useWorkbookShellLayout,
} from "../use-workbook-shell-layout.js";

function ShellLayoutHarness(props: { documentId: string }) {
  const layout = useWorkbookShellLayout({
    documentId: props.documentId,
    availableTabs: ["assistant", "changes"],
  });

  return (
    <div
      data-open={String(layout.isSideRailOpen)}
      data-tab={layout.activeSideRailTab ?? ""}
      data-testid="shell-layout-state"
      data-width={String(layout.sideRailWidth)}
    >
      <button
        data-testid="toggle-assistant"
        type="button"
        onClick={() => layout.toggleSideRail("assistant")}
      />
      <button
        data-testid="toggle-changes"
        type="button"
        onClick={() => layout.toggleSideRail("changes")}
      />
      <button data-testid="set-width" type="button" onClick={() => layout.setSideRailWidth(416)} />
    </div>
  );
}

describe("workbook shell layout", () => {
  beforeEach(() => {
    const backingStore = new Map<string, string>();
    const storage = {
      clear() {
        backingStore.clear();
      },
      getItem(key: string) {
        return backingStore.get(key) ?? null;
      },
      key(index: number) {
        return [...backingStore.keys()][index] ?? null;
      },
      removeItem(key: string) {
        backingStore.delete(key);
      },
      setItem(key: string, value: string) {
        backingStore.set(key, value);
      },
      get length() {
        return backingStore.size;
      },
    };
    Object.defineProperty(window, "localStorage", {
      configurable: true,
      value: storage,
    });
  });

  afterEach(() => {
    document.body.innerHTML = "";
  });

  it("defaults to a collapsed side rail with no implicit active tab", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<ShellLayoutHarness documentId="doc-1" />);
    });

    const state = host.querySelector("[data-testid='shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("false");
    expect(state?.getAttribute("data-tab")).toBe("");
    expect(state?.getAttribute("data-width")).toBe(String(DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH));

    await act(async () => {
      root.unmount();
    });
  });

  it("persists the active tab, open state, and side rail width across remounts", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);

    const firstRoot = createRoot(host);
    await act(async () => {
      firstRoot.render(<ShellLayoutHarness documentId="doc-2" />);
    });

    await act(async () => {
      host
        .querySelector("[data-testid='toggle-changes']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      host
        .querySelector("[data-testid='set-width']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      firstRoot.unmount();
    });

    const secondRoot = createRoot(host);
    await act(async () => {
      secondRoot.render(<ShellLayoutHarness documentId="doc-2" />);
    });

    const state = host.querySelector("[data-testid='shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("true");
    expect(state?.getAttribute("data-tab")).toBe("changes");
    expect(state?.getAttribute("data-width")).toBe("416");

    await act(async () => {
      secondRoot.unmount();
    });
  });

  it("clamps oversized persisted rail widths back into the supported range", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    window.localStorage.setItem(
      "bilig:workbook-shell-layout:doc-oversized",
      JSON.stringify({
        sideRailOpen: true,
        sideRailTab: "assistant",
        sideRailWidth: 640,
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<ShellLayoutHarness documentId="doc-oversized" />);
    });

    const state = host.querySelector("[data-testid='shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("true");
    expect(state?.getAttribute("data-tab")).toBe("assistant");
    expect(state?.getAttribute("data-width")).toBe("420");

    await act(async () => {
      root.unmount();
    });
  });

  it("uses a viewport-aware width clamp on narrow windows", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const previousInnerWidth = window.innerWidth;
    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      value: 720,
    });

    window.localStorage.setItem(
      "bilig:workbook-shell-layout:doc-narrow",
      JSON.stringify({
        sideRailOpen: true,
        sideRailTab: "assistant",
        sideRailWidth: 420,
      }),
    );

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<ShellLayoutHarness documentId="doc-narrow" />);
    });

    const state = host.querySelector("[data-testid='shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("true");
    expect(state?.getAttribute("data-tab")).toBe("assistant");
    expect(state?.getAttribute("data-width")).toBe("302");

    await act(async () => {
      root.unmount();
    });

    Object.defineProperty(window, "innerWidth", {
      configurable: true,
      value: previousInnerWidth,
    });
  });

  it("closes the rail when the active tab toggle is pressed again", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<ShellLayoutHarness documentId="doc-3" />);
    });

    await act(async () => {
      host
        .querySelector("[data-testid='toggle-assistant']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      host
        .querySelector("[data-testid='toggle-assistant']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    const state = host.querySelector("[data-testid='shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("false");
    expect(state?.getAttribute("data-tab")).toBe("assistant");

    await act(async () => {
      root.unmount();
    });
  });

  it("isolates persisted layout state per explicit persistence key", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    function ScopedShellLayoutHarness(props: { documentId: string; persistenceKey: string }) {
      const layout = useWorkbookShellLayout({
        documentId: props.documentId,
        persistenceKey: props.persistenceKey,
        availableTabs: ["assistant", "changes"],
      });

      return (
        <div
          data-open={String(layout.isSideRailOpen)}
          data-tab={layout.activeSideRailTab ?? ""}
          data-testid="scoped-shell-layout-state"
          data-width={String(layout.sideRailWidth)}
        >
          <button
            data-testid="scoped-toggle-assistant"
            type="button"
            onClick={() => layout.toggleSideRail("assistant")}
          />
          <button
            data-testid="scoped-set-width"
            type="button"
            onClick={() => layout.setSideRailWidth(432)}
          />
        </div>
      );
    }

    const host = document.createElement("div");
    document.body.appendChild(host);

    const firstRoot = createRoot(host);
    await act(async () => {
      firstRoot.render(
        <ScopedShellLayoutHarness documentId="doc-4" persistenceKey="doc-4:user-a" />,
      );
    });

    await act(async () => {
      host
        .querySelector("[data-testid='scoped-toggle-assistant']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
      host
        .querySelector("[data-testid='scoped-set-width']")
        ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    await act(async () => {
      firstRoot.unmount();
    });

    const secondRoot = createRoot(host);
    await act(async () => {
      secondRoot.render(
        <ScopedShellLayoutHarness documentId="doc-4" persistenceKey="doc-4:user-b" />,
      );
    });

    const state = host.querySelector("[data-testid='scoped-shell-layout-state']");
    expect(state?.getAttribute("data-open")).toBe("false");
    expect(state?.getAttribute("data-tab")).toBe("");
    expect(state?.getAttribute("data-width")).toBe(String(DEFAULT_WORKBOOK_SIDE_RAIL_WIDTH));

    await act(async () => {
      secondRoot.unmount();
    });
  });
});
