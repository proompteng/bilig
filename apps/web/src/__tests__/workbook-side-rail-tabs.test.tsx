// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WorkbookSideRailTabs } from "../WorkbookSideRailTabs.js";

afterEach(() => {
  document.body.innerHTML = "";
});

describe("workbook side rail tabs", () => {
  it("renders Base UI tabs with count badges and switches the active tab", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookSideRailTabs
          defaultValue="assistant"
          tabs={[
            {
              value: "assistant",
              label: "Assistant",
              panel: <div data-testid="assistant-panel">Assistant panel</div>,
            },
            {
              value: "changes",
              label: "Changes",
              count: 2,
              panel: <div data-testid="changes-panel">Changes panel</div>,
            },
          ]}
        />,
      );
    });

    const assistantTab = host.querySelector("[data-testid='workbook-side-rail-tab-assistant']");
    const changesTab = host.querySelector("[data-testid='workbook-side-rail-tab-changes']");

    expect(assistantTab?.getAttribute("aria-selected")).toBe("true");
    expect(assistantTab?.className).toContain("font-semibold");
    expect(changesTab?.textContent).toContain("2");

    await act(async () => {
      changesTab?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(changesTab?.getAttribute("aria-selected")).toBe("true");
    expect(changesTab?.className).toContain("font-semibold");
    expect(host.querySelector("[data-testid='workbook-side-rail-panel-changes']")).not.toBeNull();

    await act(async () => {
      root.unmount();
    });
  });

  it("supports a controlled active tab", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const onValueChange = vi.fn();
    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <WorkbookSideRailTabs
          tabs={[
            {
              value: "assistant",
              label: "Assistant",
              panel: <div>Assistant panel</div>,
            },
            {
              value: "changes",
              label: "Changes",
              panel: <div>Changes panel</div>,
            },
          ]}
          value="changes"
          onValueChange={onValueChange}
        />,
      );
    });

    const assistantTab = host.querySelector("[data-testid='workbook-side-rail-tab-assistant']");
    const changesTab = host.querySelector("[data-testid='workbook-side-rail-tab-changes']");

    expect(changesTab?.getAttribute("aria-selected")).toBe("true");

    await act(async () => {
      assistantTab?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });

    expect(onValueChange).toHaveBeenCalledWith("assistant");
    expect(changesTab?.getAttribute("aria-selected")).toBe("true");

    await act(async () => {
      root.unmount();
    });
  });
});
