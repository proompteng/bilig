import { Button } from "@base-ui/react/button";
import { cva } from "class-variance-authority";
import { cn } from "./cn.js";
import {
  formatWorkbookChangeTimestamp,
  type WorkbookChangeEntry,
} from "./workbook-changes-model.js";
import {
  workbookButtonClass,
  workbookPillClass,
  workbookSurfaceClass,
} from "./workbook-shell-chrome.js";

const changeEventToneClass = cva("", {
  variants: {
    tone: {
      neutral: workbookPillClass({ tone: "neutral", weight: "strong" }),
      danger: workbookPillClass({ tone: "danger", weight: "strong" }),
    },
  },
});

const CHANGE_EVENT_TONES: Record<string, "neutral" | "danger"> = {
  setCellValue: "neutral",
  setCellFormula: "neutral",
  clearCell: "danger",
  clearRange: "danger",
  fillRange: "neutral",
  copyRange: "neutral",
  moveRange: "neutral",
  updateRowMetadata: "neutral",
  updateColumnMetadata: "neutral",
  updateColumnWidth: "neutral",
  setFreezePane: "neutral",
  setRangeStyle: "neutral",
  clearRangeStyle: "neutral",
  setRangeNumberFormat: "neutral",
  clearRangeNumberFormat: "neutral",
  renderCommit: "neutral",
  restoreVersion: "neutral",
  revertChange: "danger",
  redoChange: "neutral",
  applyBatch: "neutral",
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
    case "updateRowMetadata":
    case "updateColumnMetadata":
    case "updateColumnWidth":
    case "setFreezePane":
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
    case "revertChange":
      return "Revert";
    case "redoChange":
      return "Redo";
    case "applyBatch":
      return "Sync";
    default:
      return "Change";
  }
}

function renderChangeStatus(change: WorkbookChangeEntry): string | null {
  if (change.revertedByRevision !== null) {
    return `Reverted in r${change.revertedByRevision}`;
  }
  if (change.revertsRevision !== null) {
    return `Reverted r${change.revertsRevision}`;
  }
  return null;
}

function WorkbookChangeRow(props: {
  readonly change: WorkbookChangeEntry;
  readonly isPending: boolean;
  readonly onJump: (sheetName: string, address: string) => void;
  readonly onRevert: (change: WorkbookChangeEntry) => void;
}) {
  const { change, isPending } = props;
  const statusLabel = renderChangeStatus(change);
  const content = (
    <>
      <div className="flex min-w-0 flex-wrap items-center gap-1.5">
        <span
          className={cn(
            changeEventToneClass({
              tone: CHANGE_EVENT_TONES[change.eventKind] ?? "neutral",
            }),
          )}
        >
          {formatEventLabel(change.eventKind)}
        </span>
        <span
          className={workbookPillClass({ tone: "neutral" })}
          data-testid="workbook-change-revision"
        >
          r{change.revision}
        </span>
        {statusLabel ? (
          <span
            className={workbookPillClass({ tone: "neutral", weight: "strong" })}
            data-testid="workbook-change-status"
          >
            {statusLabel}
          </span>
        ) : null}
      </div>
      <div className="mt-1.5 text-[13px] font-medium leading-5 text-[var(--wb-text)]">
        {change.summary}
      </div>
      <div className="mt-1 flex flex-wrap items-center gap-x-1.5 gap-y-1 text-[11px] text-[var(--wb-text-subtle)]">
        <span>{change.actorLabel}</span>
        <span aria-hidden="true">•</span>
        <span>{formatWorkbookChangeTimestamp(change.createdAt)}</span>
        {change.targetLabel ? (
          <>
            <span aria-hidden="true">•</span>
            <span data-testid="workbook-change-target">{change.targetLabel}</span>
          </>
        ) : null}
      </div>
    </>
  );

  const contentRegion = change.isJumpable ? (
    <Button
      className="min-w-0 flex-1 rounded-[var(--wb-radius-control)] p-0 text-left text-inherit transition-colors hover:bg-[var(--wb-hover)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
      type="button"
      onClick={() => {
        if (change.sheetName && change.address) {
          props.onJump(change.sheetName, change.address);
        }
      }}
    >
      <div className="px-2 py-1.5">{content}</div>
    </Button>
  ) : (
    <div className="min-w-0 flex-1 px-2 py-1.5 opacity-90">{content}</div>
  );

  return (
    <div
      className={cn(workbookSurfaceClass({ emphasis: "flat" }), "px-2.5 py-2")}
      data-testid="workbook-change-row"
    >
      <div className="flex items-start gap-2">
        {contentRegion}
        {change.canRevert ? (
          <Button
            className={workbookButtonClass({ tone: "danger", size: "sm" })}
            data-testid="workbook-change-revert"
            disabled={isPending}
            type="button"
            onClick={() => {
              props.onRevert(change);
            }}
          >
            {isPending ? "Reverting..." : "Revert"}
          </Button>
        ) : null}
      </div>
    </div>
  );
}

export function WorkbookChangesPanel(props: {
  readonly changes: readonly WorkbookChangeEntry[];
  readonly onJump: (sheetName: string, address: string) => void;
  readonly onRevert: (change: WorkbookChangeEntry) => void;
  readonly pendingRevisions: readonly number[];
}) {
  return (
    <div
      aria-label="Workbook changes"
      className="flex h-full min-h-0 flex-col"
      data-testid="workbook-changes-panel"
      id="workbook-changes-panel"
    >
      <div className="min-h-0 flex-1 overflow-y-auto px-3 py-3">
        {props.changes.length === 0 ? (
          <div />
        ) : (
          <div className="flex flex-col gap-1">
            {props.changes.map((change) => (
              <WorkbookChangeRow
                key={`${change.revision}:${change.summary}`}
                change={change}
                isPending={props.pendingRevisions.includes(change.revision)}
                onJump={props.onJump}
                onRevert={props.onRevert}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
