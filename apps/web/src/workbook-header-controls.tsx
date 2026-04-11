import type { ReactNode } from "react";
import { cn } from "./cn.js";
import { workbookStatusDotClass } from "./workbook-shell-chrome.js";

const HEADER_ACTION_CLASS =
  "inline-flex h-full items-center gap-2 rounded-[calc(var(--wb-radius-control)-2px)] px-2.5 text-[12px] font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1";

export function WorkbookHeaderActionButton(props: {
  children: ReactNode;
  "aria-controls"?: string;
  "aria-expanded"?: boolean;
  "aria-label"?: string;
  "data-testid"?: string;
  disabled?: boolean;
  isActive?: boolean;
  isGrouped?: boolean;
  onClick: () => void;
  type?: "button" | "submit" | "reset";
}) {
  const grouped = props.isGrouped ?? false;
  return (
    <button
      aria-controls={props["aria-controls"]}
      aria-expanded={props["aria-expanded"]}
      aria-label={props["aria-label"]}
      className={cn(
        HEADER_ACTION_CLASS,
        grouped
          ? props.isActive
            ? "bg-[var(--wb-surface)] text-[var(--wb-text)]"
            : "bg-transparent text-[var(--wb-text-muted)] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)]"
          : props.isActive
            ? "border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)]"
            : "border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)]",
        props.disabled ? "cursor-not-allowed opacity-50" : null,
      )}
      data-testid={props["data-testid"]}
      disabled={props.disabled}
      type={props.type ?? "button"}
      onClick={props.onClick}
    >
      {props.children}
    </button>
  );
}

export function WorkbookHeaderStatusChip(props: { modeLabel: string; syncLabel: string }) {
  const toneClass =
    props.syncLabel === "Ready"
      ? workbookStatusDotClass({ tone: "ready" })
      : props.syncLabel === "Syncing" || props.syncLabel === "Loading"
        ? workbookStatusDotClass({ tone: "pending" })
        : workbookStatusDotClass({ tone: "danger" });

  return (
    <>
      <span
        aria-label={`Workbook status: ${props.modeLabel}, ${props.syncLabel}`}
        className={cn(
          "inline-flex h-8 items-center gap-2 rounded-md border border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 text-[12px] font-medium text-[var(--color-mauve-700)] shadow-[0_1px_2px_rgba(15,23,42,0.04)]",
        )}
        data-testid="status-mode"
        role="status"
        title={`${props.modeLabel} • ${props.syncLabel}`}
      >
        <span aria-hidden="true" className={toneClass} />
        <span className="text-[var(--color-mauve-900)]">{props.syncLabel}</span>
      </span>
      <span className="sr-only" data-testid="status-sync">
        {props.syncLabel}
      </span>
    </>
  );
}
