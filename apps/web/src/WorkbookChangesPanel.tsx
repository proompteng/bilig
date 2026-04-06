import { cn } from "./cn.js";
import {
  formatWorkbookChangeTimestamp,
  type WorkbookChangeEntry,
} from "./workbook-changes-model.js";

const CHANGE_EVENT_TONE_CLASS_NAMES: Record<string, string> = {
  setCellValue: "bg-[#e0f2fe] text-[#075985]",
  setCellFormula: "bg-[#e0f2fe] text-[#075985]",
  clearCell: "bg-[#fee2e2] text-[#991b1b]",
  clearRange: "bg-[#fee2e2] text-[#991b1b]",
  fillRange: "bg-[#dcfce7] text-[#166534]",
  copyRange: "bg-[#dcfce7] text-[#166534]",
  moveRange: "bg-[#dcfce7] text-[#166534]",
  updateColumnWidth: "bg-[#ede9fe] text-[#5b21b6]",
  setRangeStyle: "bg-[#fef3c7] text-[#92400e]",
  clearRangeStyle: "bg-[#fef3c7] text-[#92400e]",
  setRangeNumberFormat: "bg-[#fef3c7] text-[#92400e]",
  clearRangeNumberFormat: "bg-[#fef3c7] text-[#92400e]",
  renderCommit: "bg-[#e2e8f0] text-[#334155]",
  restoreVersion: "bg-[#dbeafe] text-[#1d4ed8]",
  applyBatch: "bg-[#e2e8f0] text-[#334155]",
};

function formatEventLabel(eventKind: string): string {
  switch (eventKind) {
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
    case "clearRange":
      return "Value";
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return "Range";
    case "updateColumnWidth":
      return "Layout";
    case "setRangeStyle":
    case "clearRangeStyle":
    case "setRangeNumberFormat":
    case "clearRangeNumberFormat":
      return "Format";
    case "renderCommit":
      return "Batch";
    case "restoreVersion":
      return "Version";
    case "applyBatch":
      return "Sync";
    default:
      return "Change";
  }
}

export function WorkbookChangesPanel(props: {
  readonly changes: readonly WorkbookChangeEntry[];
  readonly isOpen: boolean;
  readonly onClose: () => void;
  readonly onJump: (sheetName: string, address: string) => void;
}) {
  if (!props.isOpen) {
    return null;
  }

  return (
    <aside
      aria-label="Workbook changes"
      className="absolute inset-y-3 right-3 z-20 flex w-[22rem] flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[0_20px_48px_rgba(15,23,42,0.16)]"
      data-testid="workbook-changes-panel"
      id="workbook-changes-panel"
    >
      <div className="flex items-center justify-between gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-3">
        <div className="min-w-0">
          <h2 className="text-[13px] font-semibold text-[var(--wb-text)]">Changes</h2>
          <p className="text-[11px] text-[var(--wb-text-subtle)]">
            {props.changes.length === 0
              ? "No authoritative changes yet"
              : `${props.changes.length} recent authoritative changes`}
          </p>
        </div>
        <button
          aria-label="Close changes pane"
          className="inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          <span aria-hidden="true">×</span>
        </button>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.changes.length === 0 ? (
          <div className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-4 py-5 text-sm text-[var(--wb-text-subtle)]">
            New workbook edits will appear here with jump targets into the affected sheet.
          </div>
        ) : (
          <div className="flex flex-col gap-2">
            {props.changes.map((change) => (
              <button
                key={`${change.revision}:${change.summary}`}
                className={cn(
                  "rounded-[var(--wb-radius-control)] border px-3 py-3 text-left transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1",
                  change.isJumpable
                    ? "border-[var(--wb-border)] bg-[var(--wb-surface)] hover:border-[var(--wb-accent-ring)] hover:bg-[var(--wb-surface-subtle)]"
                    : "cursor-default border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] opacity-80",
                )}
                data-testid="workbook-change-row"
                disabled={!change.isJumpable}
                type="button"
                onClick={() => {
                  if (change.isJumpable && change.sheetName && change.address) {
                    props.onJump(change.sheetName, change.address);
                  }
                }}
              >
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="flex items-center gap-2">
                      <span
                        className={cn(
                          "inline-flex h-5 items-center rounded-full px-2 text-[10px] font-semibold uppercase tracking-[0.04em]",
                          CHANGE_EVENT_TONE_CLASS_NAMES[change.eventKind] ??
                            "bg-[#e2e8f0] text-[#334155]",
                        )}
                      >
                        {formatEventLabel(change.eventKind)}
                      </span>
                      <span className="text-[11px] text-[var(--wb-text-subtle)]">
                        r{change.revision}
                      </span>
                    </div>
                    <div className="mt-2 text-[13px] font-medium leading-5 text-[var(--wb-text)]">
                      {change.summary}
                    </div>
                    <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
                      <span>{change.actorLabel}</span>
                      <span aria-hidden="true">•</span>
                      <span>{formatWorkbookChangeTimestamp(change.createdAt)}</span>
                    </div>
                  </div>
                </div>
                <div className="mt-2 text-[11px] text-[var(--wb-text-subtle)]">
                  {change.targetLabel ? (
                    change.isJumpable ? (
                      <>Jump to {change.targetLabel}</>
                    ) : (
                      <>{change.targetLabel} is no longer available</>
                    )
                  ) : (
                    <>No jump target</>
                  )}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </aside>
  );
}
