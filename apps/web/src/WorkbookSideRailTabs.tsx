import { useEffect, useMemo, useState, type ReactNode } from "react";
import { Tabs } from "@base-ui/react/tabs";
import { cva } from "class-variance-authority";
import { cn } from "./cn.js";

export interface WorkbookSideRailTabDefinition {
  readonly value: string;
  readonly label: string;
  readonly panel: ReactNode;
  readonly count?: number | undefined;
}

const railRootClass = cva(
  "flex h-full min-h-0 w-full flex-col overflow-hidden bg-[var(--wb-surface)]",
);

const railListClass = cva(
  "relative z-0 flex gap-1 border-b border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-2",
);

const railTabClass = cva(
  "group flex h-8 items-center justify-center gap-1.5 border-0 px-3 text-[12px] font-medium break-keep whitespace-nowrap outline-none select-none transition-colors before:inset-x-0 before:inset-y-1 before:rounded-[calc(var(--wb-radius-control)-2px)] before:-outline-offset-1 before:outline-[var(--wb-accent)] hover:text-[var(--wb-text)] focus-visible:relative focus-visible:before:absolute focus-visible:before:outline focus-visible:before:outline-2",
  {
    variants: {
      active: {
        true: "font-semibold text-[var(--wb-text)]",
        false: "text-[var(--wb-text-muted)]",
      },
    },
  },
);

const railIndicatorClass = cva(
  "absolute top-1/2 left-0 z-[-1] h-8 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] -translate-y-1/2 rounded-[calc(var(--wb-radius-control)-2px)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)] transition-[translate,width] duration-200 ease-out",
);

const railPanelClass = cva("min-h-0 flex-1 overflow-hidden bg-[var(--wb-app-bg)]");

const railCountBadgeClass = cva(
  "inline-flex min-w-4 items-center justify-center rounded-full px-1.5 text-[10px] font-semibold leading-none",
  {
    variants: {
      active: {
        true: "bg-[var(--wb-accent-soft)] text-[var(--wb-accent)]",
        false: "bg-[var(--wb-surface-subtle)] text-[var(--wb-text-subtle)]",
      },
    },
  },
);

export function WorkbookSideRailTabs(props: {
  readonly tabs: readonly WorkbookSideRailTabDefinition[];
  readonly defaultValue?: string;
}) {
  const tabs = useMemo(() => props.tabs.filter((tab) => tab.panel != null), [props.tabs]);
  const resolvedDefaultValue = useMemo(() => {
    if (tabs.length === 0) {
      return null;
    }
    return tabs.some((tab) => tab.value === props.defaultValue)
      ? (props.defaultValue ?? tabs[0]!.value)
      : tabs[0]!.value;
  }, [props.defaultValue, tabs]);
  const [value, setValue] = useState<string | null>(resolvedDefaultValue);

  useEffect(() => {
    if (resolvedDefaultValue === null) {
      setValue(null);
      return;
    }
    if (!tabs.some((tab) => tab.value === value)) {
      setValue(resolvedDefaultValue);
    }
  }, [resolvedDefaultValue, tabs, value]);

  if (value === null || tabs.length === 0) {
    return null;
  }

  return (
    <Tabs.Root
      className={railRootClass()}
      value={value}
      onValueChange={(nextValue) => {
        setValue(String(nextValue));
      }}
    >
      <Tabs.List aria-label="Workbook panels" className={railListClass()}>
        {tabs.map((tab) => (
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
                  railCountBadgeClass({
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
      {tabs.map((tab) => (
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
