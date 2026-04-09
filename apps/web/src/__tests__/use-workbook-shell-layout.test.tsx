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
    defaultTab: "assistant",
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
    window.localStorage.clear();
  });

  afterEach(() => {
    document.body.innerHTML = "";
    window.localStorage.clear();
  });

  it("defaults to a collapsed side rail with the assistant remembered as the fallback tab", async () => {
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
    expect(state?.getAttribute("data-tab")).toBe("assistant");
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
});
