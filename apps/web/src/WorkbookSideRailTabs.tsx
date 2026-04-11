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
  "flex h-full min-h-0 w-full flex-col overflow-hidden bg-[var(--color-mauve-50)]",
);

const railListClass = cva(
  "relative flex items-end gap-1 border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-3 pt-2",
);

const railTabClass = cva(
  "group relative inline-flex h-9 items-center justify-center gap-1.5 rounded-t-md border-b-2 border-transparent px-3 pb-2 text-[13px] font-medium break-keep whitespace-nowrap outline-none select-none transition-[color,background-color] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]",
  {
    variants: {
      active: {
        true: "bg-[var(--color-mauve-50)] font-semibold text-[var(--color-mauve-950)]",
        false:
          "bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-100)]/70 hover:text-[var(--color-mauve-900)]",
      },
    },
  },
);

const railIndicatorClass = cva(
  "absolute bottom-0 left-0 h-0.5 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] rounded-full bg-[var(--color-mauve-700)] transition-[translate,width] duration-200 ease-out",
);

const railPanelClass = cva("min-h-0 flex-1 overflow-hidden bg-[var(--color-mauve-50)]");

const railCountClass = cva(
  "inline-flex items-center justify-center text-[11px] font-semibold tabular-nums leading-none transition-colors",
  {
    variants: {
      active: {
        true: "text-[var(--color-mauve-700)]",
        false: "text-[var(--color-mauve-500)] group-hover:text-[var(--color-mauve-700)]",
      },
    },
  },
);

export function WorkbookSideRailTabs(props: {
  readonly tabs: readonly WorkbookSideRailTabDefinition[];
  readonly defaultValue?: string;
  readonly value?: string | null;
  readonly onValueChange?: ((nextValue: string) => void) | undefined;
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
  const isControlled = props.value !== undefined;
  const [uncontrolledValue, setUncontrolledValue] = useState<string | null>(resolvedDefaultValue);
  const value = isControlled ? (props.value ?? null) : uncontrolledValue;

  useEffect(() => {
    if (isControlled) {
      return;
    }
    if (resolvedDefaultValue === null) {
      setUncontrolledValue(null);
      return;
    }
    if (!tabs.some((tab) => tab.value === value)) {
      setUncontrolledValue(resolvedDefaultValue);
    }
  }, [isControlled, resolvedDefaultValue, tabs, value]);

  if (value === null || tabs.length === 0) {
    return null;
  }

  return (
    <Tabs.Root
      className={railRootClass()}
      value={value}
      onValueChange={(nextValue) => {
        const resolvedNextValue = String(nextValue);
        if (!isControlled) {
          setUncontrolledValue(resolvedNextValue);
        }
        props.onValueChange?.(resolvedNextValue);
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
