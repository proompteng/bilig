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
  "relative z-0 flex gap-1 border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 py-1.5",
);

const railTabClass = cva(
  "group relative flex h-8 items-center justify-center gap-1.5 rounded-md border px-3 text-[12px] font-medium break-keep whitespace-nowrap outline-none select-none transition-[background-color,border-color,color,box-shadow] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]",
  {
    variants: {
      active: {
        true: "border-[var(--color-mauve-300)] bg-white font-semibold text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
        false:
          "border-transparent bg-transparent text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]",
      },
    },
  },
);

const railIndicatorClass = cva(
  "absolute top-1/2 left-0 z-[-1] h-8 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] -translate-y-1/2 rounded-md border border-[var(--color-mauve-300)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)] transition-[translate,width] duration-200 ease-out",
);

const railPanelClass = cva("min-h-0 flex-1 overflow-hidden bg-[var(--color-mauve-50)]");

const railCountBadgeClass = cva(
  "inline-flex min-w-4 items-center justify-center rounded-full px-1.5 text-[10px] font-semibold leading-none transition-colors",
  {
    variants: {
      active: {
        true: "bg-[var(--color-mauve-100)] text-[var(--color-mauve-900)]",
        false: "bg-[var(--color-mauve-100)] text-[var(--color-mauve-700)]",
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
