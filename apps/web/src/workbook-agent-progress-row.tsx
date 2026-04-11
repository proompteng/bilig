import { Fragment, useEffect, useMemo, useState } from "react";
import { cn } from "./cn.js";
import {
  workbookInsetClass,
  workbookStatusDotClass,
  workbookSurfaceClass,
} from "./workbook-shell-chrome.js";

const progressStages = [
  { key: "selection", label: "Reading selection", shortLabel: "Selection" },
  { key: "context", label: "Checking context", shortLabel: "Context" },
  { key: "draft", label: "Drafting reply", shortLabel: "Draft" },
] as const;
const thinkingLetters = [
  { key: "t", label: "T" },
  { key: "h", label: "h" },
  { key: "i-1", label: "i" },
  { key: "n-1", label: "n" },
  { key: "k", label: "k" },
  { key: "i-2", label: "i" },
  { key: "n-2", label: "n" },
  { key: "g", label: "g" },
] as const;

const PROGRESS_STAGE_INTERVAL_MS = 1400;

function ThinkingWave() {
  return (
    <span
      aria-label="Thinking"
      className="inline-flex items-center gap-[1px] text-[11px] font-medium text-[var(--wb-text-muted)]"
    >
      {thinkingLetters.map((letter, index) => (
        <span
          key={letter.key}
          aria-hidden="true"
          className="workbook-agent-thinking-letter"
          style={{ animationDelay: `${String(index * 90)}ms` }}
        >
          {letter.label}
        </span>
      ))}
    </span>
  );
}

export function WorkbookAgentProgressRow() {
  const [stageIndex, setStageIndex] = useState(0);

  useEffect(() => {
    const interval = window.setInterval(() => {
      setStageIndex((currentIndex) =>
        currentIndex < progressStages.length - 1 ? currentIndex + 1 : currentIndex,
      );
    }, PROGRESS_STAGE_INTERVAL_MS);
    return () => {
      window.clearInterval(interval);
    };
  }, []);

  const currentStage = progressStages[stageIndex] ?? progressStages[0];
  const progressSummary = useMemo(
    () =>
      progressStages.map((stage, index) => {
        const state =
          index < stageIndex ? "complete" : index === stageIndex ? "active" : "upcoming";
        return {
          key: stage.key,
          label: stage.label,
          shortLabel: stage.shortLabel,
          state,
        };
      }),
    [stageIndex],
  );

  return (
    <div className="px-1 py-1.5">
      <div
        aria-live="polite"
        className={cn(workbookSurfaceClass({ emphasis: "flat" }), "px-2.5 py-2")}
        data-testid="workbook-agent-progress-row"
        role="status"
      >
        <div className="flex items-center gap-2">
          <div className="inline-flex items-center gap-1.5">
            <span aria-hidden="true" className={workbookStatusDotClass({ tone: "pending" })} />
            <ThinkingWave />
          </div>
          <div className="h-px flex-1 bg-[var(--wb-border)]" />
          <div className="text-[12px] font-medium text-[var(--wb-text)]">{currentStage.label}</div>
        </div>
        <div className="mt-1.5 flex flex-wrap items-center gap-1.5 text-[11px] leading-none">
          {progressSummary.map((stage, index) => (
            <Fragment key={stage.key}>
              {index > 0 ? (
                <span aria-hidden="true" className="text-[var(--wb-text-subtle)]">
                  ·
                </span>
              ) : null}
              <span
                className={cn(
                  "transition-colors",
                  stage.state === "active"
                    ? "font-medium text-[var(--wb-text)]"
                    : stage.state === "complete"
                      ? "text-[var(--wb-text-muted)]"
                      : "text-[var(--wb-text-subtle)]",
                )}
              >
                {stage.shortLabel}
              </span>
            </Fragment>
          ))}
        </div>
        <div
          aria-hidden="true"
          className={cn(workbookInsetClass(), "mt-2 grid grid-cols-3 gap-1 overflow-hidden p-1")}
        >
          {progressSummary.map((stage) => (
            <span
              key={stage.key}
              className={cn(
                "h-1 rounded-full transition-colors",
                stage.state === "active"
                  ? "bg-[var(--wb-accent)]"
                  : stage.state === "complete"
                    ? "bg-[var(--wb-border-strong)]"
                    : "bg-[var(--wb-border)]",
              )}
            />
          ))}
        </div>
      </div>
    </div>
  );
}
