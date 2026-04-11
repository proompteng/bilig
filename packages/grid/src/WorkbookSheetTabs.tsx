import React from "react";
import { cva } from "class-variance-authority";

interface WorkbookSheetTabsProps {
  sheetName: string;
  sheetNames: string[];
  selectionStatus?: React.ReactNode;
  onSelectSheet(this: void, sheetName: string): void;
  onCreateSheet?: (() => void) | undefined;
  onRenameSheet?: ((currentName: string, nextName: string) => void) | undefined;
}

const sheetStripClass = cva(
  "flex min-h-12 items-center justify-between gap-3 border-t border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] px-2.5 py-1.5",
);

const sheetTabsShellClass = cva("flex min-w-0 flex-1 items-center gap-2 overflow-hidden");

const sheetListClass = cva(
  "inline-flex min-w-0 flex-1 items-center gap-1 overflow-x-auto rounded-lg border border-[var(--color-mauve-200)] bg-[var(--color-mauve-100)] p-1 shadow-[inset_0_1px_0_rgba(255,255,255,0.85)]",
);

const sheetRenameShellClass = cva(
  "inline-flex h-8 items-center rounded-md border border-[var(--color-mauve-300)] bg-white px-3 shadow-[0_1px_2px_rgba(15,23,42,0.06)]",
);

const sheetActionButtonClass = cva(
  "inline-flex size-8 shrink-0 items-center justify-center rounded-md border border-[var(--color-mauve-200)] bg-[var(--color-mauve-100)] text-[var(--color-mauve-700)] outline-none transition-[background-color,border-color,color,box-shadow] hover:bg-[var(--color-mauve-200)] hover:text-[var(--color-mauve-900)] focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-50)] disabled:cursor-not-allowed disabled:opacity-50",
);

const sheetTabClass = cva(
  "inline-flex h-8 items-center rounded-md border px-3.5 text-[12px] whitespace-nowrap outline-none transition-[background-color,border-color,color,box-shadow] duration-150 focus-visible:ring-2 focus-visible:ring-[var(--color-mauve-400)] focus-visible:ring-offset-1 focus-visible:ring-offset-[var(--color-mauve-100)]",
  {
    variants: {
      active: {
        true:
          "border-[var(--color-mauve-300)] bg-white font-semibold text-[var(--color-mauve-900)] shadow-[0_1px_2px_rgba(15,23,42,0.06)]",
        false:
          "border-transparent bg-transparent font-medium text-[var(--color-mauve-600)] hover:bg-[var(--color-mauve-200)] hover:text-[var(--color-mauve-900)]",
      },
      disabled: {
        true: "cursor-not-allowed opacity-50",
        false: "",
      },
    },
  },
);

const sheetStatusSlotClass = cva(
  "inline-flex min-h-10 items-center rounded-lg border border-[var(--color-mauve-200)] bg-[var(--color-mauve-100)] px-1.5 py-1 shadow-[inset_0_1px_0_rgba(255,255,255,0.85)]",
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
  const tabRefs = React.useRef<Record<string, HTMLButtonElement | null>>({});

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

  const focusSheetTab = React.useCallback((targetSheetName: string) => {
    tabRefs.current[targetSheetName]?.focus();
  }, []);

  const moveSheetSelection = React.useCallback(
    (targetIndex: number) => {
      const resolvedIndex = Math.max(0, Math.min(sheetNames.length - 1, targetIndex));
      const targetSheetName = sheetNames[resolvedIndex];
      if (!targetSheetName) {
        return;
      }
      onSelectSheet(targetSheetName);
      queueMicrotask(() => {
        focusSheetTab(targetSheetName);
      });
    },
    [focusSheetTab, onSelectSheet, sheetNames],
  );

  return (
    <div className={sheetStripClass()}>
      <div className={sheetTabsShellClass()}>
        <div aria-label="Sheets" className={sheetListClass()} role="tablist">
          {sheetNames.map((name, index) =>
            renamingSheetName === name ? (
              <div className={sheetRenameShellClass()} key={name}>
                <input
                  aria-label={`Rename ${name}`}
                  className="w-[120px] min-w-0 border-none bg-transparent p-0 text-[12px] font-semibold text-[var(--color-mauve-900)] outline-none"
                  data-testid="workbook-sheet-rename-input"
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
              <button
                aria-selected={name === sheetName}
                className={sheetTabClass({ active: name === sheetName })}
                data-testid={`workbook-sheet-tab-${name}`}
                key={name}
                onClick={() => onSelectSheet(name)}
                onDoubleClick={() => startSheetRename(name)}
                onKeyDown={(event) => {
                  if (event.key === "F2") {
                    event.preventDefault();
                    startSheetRename(name);
                    return;
                  }
                  if (event.key === "ArrowRight") {
                    event.preventDefault();
                    moveSheetSelection(index + 1);
                    return;
                  }
                  if (event.key === "ArrowLeft") {
                    event.preventDefault();
                    moveSheetSelection(index - 1);
                    return;
                  }
                  if (event.key === "Home") {
                    event.preventDefault();
                    moveSheetSelection(0);
                    return;
                  }
                  if (event.key === "End") {
                    event.preventDefault();
                    moveSheetSelection(sheetNames.length - 1);
                  }
                }}
                ref={(node) => {
                  tabRefs.current[name] = node;
                }}
                role="tab"
                tabIndex={name === sheetName ? 0 : -1}
                title={name}
                type="button"
              >
                {name}
              </button>
            ),
          )}
        </div>
        {onCreateSheet ? (
          <button
            aria-label="Create sheet"
            className={sheetActionButtonClass()}
            data-testid="workbook-sheet-add"
            onClick={onCreateSheet}
            title="Add sheet"
            type="button"
          >
            <SheetAddIcon />
          </button>
        ) : null}
      </div>
      {selectionStatus ? (
        <div className={sheetStatusSlotClass()}>
          {selectionStatus}
        </div>
      ) : null}
    </div>
  );
});
