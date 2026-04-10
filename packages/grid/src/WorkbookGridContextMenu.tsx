import type { CSSProperties, RefObject } from "react";

export interface WorkbookGridContextMenuState {
  readonly x: number;
  readonly y: number;
  readonly target: {
    readonly kind: "row" | "column";
    readonly index: number;
    readonly hidden: boolean;
  };
}

export function WorkbookGridContextMenu(props: {
  state: WorkbookGridContextMenuState;
  menuRef: RefObject<HTMLDivElement | null>;
  onClose(this: void): void;
  onToggleTargetHidden(this: void): void;
}) {
  const { state, menuRef, onClose, onToggleTargetHidden } = props;
  const targetAxisLabel = state.target.kind === "row" ? "row" : "column";
  const actionLabel = `${state.target.hidden ? "Unhide" : "Hide"} ${targetAxisLabel}`;
  const style: CSSProperties = {
    left: Math.round(state.x),
    position: "fixed",
    top: Math.round(state.y),
  };

  return (
    <div
      aria-label="Grid context menu"
      className="z-[1100] min-w-[168px] rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface-elevated)] p-1 shadow-[var(--wb-shadow-lg)]"
      data-grid-context-menu="true"
      data-testid="grid-context-menu"
      ref={menuRef}
      role="menu"
      style={style}
    >
      <button
        aria-label={actionLabel}
        className="flex h-8 w-full items-center rounded-[4px] px-2 text-left text-[11px] font-medium text-[var(--wb-text)] outline-none transition-colors hover:bg-[var(--wb-hover)] focus-visible:bg-[var(--wb-hover)]"
        data-testid={`grid-context-action-${state.target.hidden ? "unhide" : "hide"}-${targetAxisLabel}`}
        onClick={onToggleTargetHidden}
        role="menuitem"
        type="button"
      >
        {actionLabel}
      </button>
      <button
        aria-label="Close grid context menu"
        className="sr-only"
        onClick={onClose}
        type="button"
      >
        Close
      </button>
    </div>
  );
}
