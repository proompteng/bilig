import React from "react";
import { Button } from "@base-ui/react/button";
import { Tabs } from "@base-ui/react/tabs";
import { cva } from "class-variance-authority";
import { cn } from "./cn.js";

interface WorkbookSheetTabsProps {
  sheetName: string;
  sheetNames: string[];
  selectionStatus?: React.ReactNode;
  onSelectSheet(this: void, sheetName: string): void;
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
}

const sheetStripClass = cva(
  "flex min-h-11 items-center justify-between gap-3 border-t border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2.5 py-1.5",
);

const sheetTabsRootClass = cva("min-w-0 flex-1");

const sheetListClass = cva(
  "relative z-0 flex max-w-full items-center gap-1 overflow-x-auto rounded-[calc(var(--wb-radius-control)+3px)] bg-[var(--wb-surface-muted)] px-1.5 py-1.5",
);

const sheetIndicatorClass = cva(
  "pointer-events-none absolute top-1/2 left-0 z-[-1] h-8 w-[var(--active-tab-width)] translate-x-[var(--active-tab-left)] -translate-y-1/2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[var(--wb-shadow-sm)] transition-[translate,width] duration-200 ease-out",
);

const sheetRenameShellClass = cva(
  "relative z-[1] inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 shadow-[var(--wb-shadow-sm)]",
);

const sheetActionButtonClass = cva(
  "inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-[var(--wb-radius-control)] border border-transparent bg-transparent text-[var(--wb-text-subtle)] outline-none transition-[background-color,border-color,color] hover:bg-[var(--wb-surface)] hover:text-[var(--wb-text)] focus-visible:border-[var(--wb-border)] focus-visible:bg-[var(--wb-surface)] focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] disabled:cursor-not-allowed disabled:opacity-50",
);

const sheetTabClass = cva(
  "relative z-[1] flex h-8 items-center rounded-[var(--wb-radius-control)] px-4 text-[12px] outline-none transition-colors duration-150 before:inset-x-0 before:inset-y-1 before:rounded-[calc(var(--wb-radius-control)-1px)] before:-outline-offset-1 before:outline-[var(--wb-accent)] hover:text-[var(--wb-text)] focus-visible:before:absolute focus-visible:before:outline focus-visible:before:outline-2",
  {
    variants: {
      active: {
        true: "font-semibold text-[var(--wb-text)]",
        false: "font-medium text-[var(--wb-text-muted)]",
      },
      disabled: {
        true: "cursor-not-allowed opacity-50",
        false: "",
      },
    },
  },
);

function SheetAddIcon() {
  return (
    <svg aria-hidden="true" className="h-4 w-4" fill="none" viewBox="0 0 16 16">
      <path
        d="M8 3.25v9.5M3.25 8h9.5"
        stroke="currentColor"
        strokeLinecap="round"
        strokeWidth="1.5"
      />
    </svg>
  );
}

function getSheetTabClassName(state: Tabs.Tab.State): string {
  return cn(sheetTabClass({ active: state.active, disabled: state.disabled }));
}

export const WorkbookSheetTabs = React.memo(function WorkbookSheetTabs({
  sheetName,
  sheetNames,
  selectionStatus,
  onSelectSheet,
  onCreateSheet,
  onRenameSheet,
}: WorkbookSheetTabsProps) {
  const [renamingSheetName, setRenamingSheetName] = React.useState<string | null>(null);
  const [renameDraft, setRenameDraft] = React.useState("");
  const renameInputRef = React.useRef<HTMLInputElement | null>(null);

  const startSheetRename = React.useCallback(
    (targetSheetName: string) => {
      if (!onRenameSheet) {
        return;
      }
      onSelectSheet(targetSheetName);
      setRenamingSheetName(targetSheetName);
      setRenameDraft(targetSheetName);
    },
    [onRenameSheet, onSelectSheet],
  );

  const cancelSheetRename = React.useCallback(() => {
    setRenamingSheetName(null);
    setRenameDraft("");
  }, []);

  const commitSheetRename = React.useCallback(
    (targetSheetName: string) => {
      const nextName = renameDraft.trim();
      setRenamingSheetName(null);
      setRenameDraft("");
      if (!onRenameSheet || nextName.length === 0 || nextName === targetSheetName) {
        return;
      }
      onRenameSheet(targetSheetName, nextName);
    },
    [onRenameSheet, renameDraft],
  );

  React.useEffect(() => {
    if (!renamingSheetName) {
      return;
    }
    renameInputRef.current?.focus();
    renameInputRef.current?.select();
  }, [renamingSheetName]);

  React.useEffect(() => {
    if (renamingSheetName && !sheetNames.includes(renamingSheetName)) {
      cancelSheetRename();
    }
  }, [cancelSheetRename, renamingSheetName, sheetNames]);

  return (
    <div className={sheetStripClass()}>
      <div className="flex min-w-0 flex-1 items-center gap-1">
        <Tabs.Root
          className={sheetTabsRootClass()}
          value={sheetName}
          onValueChange={(value) => onSelectSheet(String(value))}
        >
          <Tabs.List aria-label="Sheets" className={sheetListClass()}>
            <Tabs.Indicator className={sheetIndicatorClass()} renderBeforeHydration />
            {sheetNames.map((name) =>
              renamingSheetName === name ? (
                <div className={sheetRenameShellClass()} key={name}>
                  <input
                    aria-label={`Rename ${name}`}
                    className="w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-semibold text-[var(--wb-text)] outline-none"
                    onBlur={() => commitSheetRename(name)}
                    onChange={(event) => setRenameDraft(event.target.value)}
                    onClick={(event) => event.stopPropagation()}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        commitSheetRename(name);
                        return;
                      }
                      if (event.key === "Escape") {
                        event.preventDefault();
                        cancelSheetRename();
                      }
                    }}
                    ref={renameInputRef}
                    value={renameDraft}
                  />
                </div>
              ) : (
                <Tabs.Tab
                  className={getSheetTabClassName}
                  key={name}
                  onDoubleClick={() => startSheetRename(name)}
                  onKeyDown={(event) => {
                    if (event.key === "F2") {
                      event.preventDefault();
                      startSheetRename(name);
                    }
                  }}
                  title={name}
                  value={name}
                >
                  {name}
                </Tabs.Tab>
              ),
            )}
          </Tabs.List>
        </Tabs.Root>
        {onCreateSheet ? (
          <Button
            aria-label="Create sheet"
            className={sheetActionButtonClass()}
            onClick={onCreateSheet}
            title="Add sheet"
            type="button"
          >
            <SheetAddIcon />
          </Button>
        ) : null}
      </div>
      {selectionStatus ? (
        <div className="inline-flex flex-wrap items-center gap-1.5">{selectionStatus}</div>
      ) : null}
    </div>
  );
});
