import type { ReactNode } from "react";
import { cva } from "class-variance-authority";

export interface WorkbookSidePanelTabDefinition {
  readonly value: string;
  readonly label: string;
  readonly panel: ReactNode;
  readonly count?: number | undefined;
}

export const panelRootClass = cva(
  "flex h-full min-h-0 w-full flex-col overflow-hidden bg-[var(--color-mauve-50)]",
);

export const panelListClass = cva(
  "relative flex min-h-11 w-full items-center gap-2 border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 py-1.5",
);

export const panelTabClass = cva(
  "group relative inline-flex h-8 items-center justify-center gap-1.5 rounded-md border-b-2 border-transparent px-2.5 text-[13px] font-medium break-keep whitespace-nowrap outline-none select-none transition-[color,background-color] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)]",
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

export const panelIndicatorClass = cva(
  "absolute bottom-0 left-0 h-0.5 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] rounded-full bg-[var(--color-mauve-700)] transition-[translate,width] duration-200 ease-out",
);

export const panelContentClass = cva("min-h-0 flex-1 overflow-hidden bg-[var(--color-mauve-50)]");

export const panelCountClass = cva(
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
