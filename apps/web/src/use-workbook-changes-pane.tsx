import { useMemo, useState } from "react";
import { WorkbookChangesPanel } from "./WorkbookChangesPanel.js";
import { useWorkbookChanges, type ZeroWorkbookChangeQuerySource } from "./use-workbook-changes.js";

export function useWorkbookChangesPane(input: {
  readonly documentId: string;
  readonly sheetNames: readonly string[];
  readonly zero: ZeroWorkbookChangeQuerySource;
  readonly enabled: boolean;
  readonly onJump: (sheetName: string, address: string) => void;
}) {
  const { documentId, enabled, onJump, sheetNames, zero } = input;
  const changes = useWorkbookChanges({
    documentId,
    sheetNames,
    zero,
    enabled,
  });
  const [isOpen, setIsOpen] = useState(false);
  const changeCount = Math.min(changes.length, 99);

  const changesToggle = useMemo(
    () => (
      <button
        aria-controls="workbook-changes-panel"
        aria-expanded={isOpen}
        aria-label={`Show workbook changes (${changes.length})`}
        className="inline-flex h-8 items-center gap-2 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
        data-testid="workbook-changes-toggle"
        type="button"
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        <span>Changes</span>
        <span className="inline-flex min-w-5 items-center justify-center rounded-full bg-[var(--wb-surface-subtle)] px-1.5 text-[11px] font-semibold text-[var(--wb-text)]">
          {changeCount}
        </span>
      </button>
    ),
    [changeCount, changes.length, isOpen],
  );

  const changesPanel = useMemo(
    () => (
      <WorkbookChangesPanel
        changes={changes}
        isOpen={isOpen}
        onClose={() => {
          setIsOpen(false);
        }}
        onJump={onJump}
      />
    ),
    [changes, isOpen, onJump],
  );

  return {
    changesPanel,
    changesToggle,
  };
}
