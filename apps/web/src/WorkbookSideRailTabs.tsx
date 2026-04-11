import type { ReactNode } from "react";
import { cva } from "class-variance-authority";

export interface WorkbookSideRailTabDefinition {
  readonly value: string;
  readonly label: string;
  readonly panel: ReactNode;
  readonly count?: number | undefined;
}

export const railRootClass = cva(
  "flex h-full min-h-0 w-full flex-col overflow-hidden bg-[var(--color-mauve-50)]",
);

export const railListClass = cva(
  "relative flex w-full items-end gap-1 border-b border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-3 pt-2",
);

export const railTabClass = cva(
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

export const railIndicatorClass = cva(
  "absolute bottom-0 left-0 h-0.5 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] rounded-full bg-[var(--color-mauve-700)] transition-[translate,width] duration-200 ease-out",
);

export const railPanelClass = cva("min-h-0 flex-1 overflow-hidden bg-[var(--color-mauve-50)]");

export const railCountClass = cva(
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
