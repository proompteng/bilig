// @vitest-environment jsdom
import { act, useState } from "react";
import { createRoot } from "react-dom/client";
import { Tabs } from "@base-ui/react/tabs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cn } from "../cn.js";
import {
  railCountClass,
  railIndicatorClass,
  railListClass,
  railPanelClass,
  railRootClass,
  railTabClass,
  type WorkbookSideRailTabDefinition,
} from "../WorkbookSideRailTabs.js";

afterEach(() => {
  document.body.innerHTML = "";
});

function SideRailHarness(props: {
  readonly defaultValue?: string;
  readonly tabs: readonly WorkbookSideRailTabDefinition[];
  readonly value?: string;
  readonly onValueChange?: (nextValue: string) => void;
}) {
  const visibleTabs = props.tabs.filter((tab) => tab.panel != null);
  const [uncontrolledValue, setUncontrolledValue] = useState<string>(
    props.defaultValue ?? visibleTabs[0]?.value ?? "",
  );
  const value = props.value ?? uncontrolledValue;

  if (!value || visibleTabs.length === 0) {
    return null;
  }

  return (
    <Tabs.Root
      className={railRootClass()}
      value={value}
      onValueChange={(nextValue) => {
        const resolvedNextValue = String(nextValue);
        if (props.value === undefined) {
          setUncontrolledValue(resolvedNextValue);
        }
        props.onValueChange?.(resolvedNextValue);
      }}
    >
      <Tabs.List aria-label="Workbook panels" className={railListClass()}>
        {visibleTabs.map((tab) => (
          <Tabs.Tab
            className={(state) => railTabClass({ active: state.active })}
            data-testid={`workbook-side-rail-tab-${tab.value}`}
            key={tab.value}
            value={tab.value}
          >
            <span>{tab.label}</span>
            {typeof tab.count === "number" ? (
              <span
                className={cn(
                  railCountClass({
                    active: value === tab.value,
                  }),
                )}
              >
                {String(Math.min(tab.count, 99))}
              </span>
            ) : null}
          </Tabs.Tab>
        ))}
        <Tabs.Indicator className={railIndicatorClass()} renderBeforeHydration />
      </Tabs.List>
      {visibleTabs.map((tab) => (
        <Tabs.Panel
          className={railPanelClass()}
          data-testid={`workbook-side-rail-panel-${tab.value}`}
          keepMounted
          key={tab.value}
          value={tab.value}
        >
          {tab.panel}
        </Tabs.Panel>
      ))}
    </Tabs.Root>
  );
}

describe("workbook side rail tabs", () => {
  it("renders Base UI tabs with count text and switches the active tab", async () => {
    (
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true;

    const host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(
        <SideRailHarness
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
    const tabList = host.querySelector("[role='tablist']");

    expect(assistantTab?.getAttribute("aria-selected")).toBe("true");
    expect(assistantTab?.className).toContain("font-semibold");
    expect(changesTab?.textContent).toContain("2");
    expect(tabList?.className).toContain("w-full");

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
        <SideRailHarness
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
