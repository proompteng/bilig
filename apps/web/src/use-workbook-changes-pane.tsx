import { useCallback, useMemo, useState } from "react";
import { mutators } from "@bilig/zero-sync";
import { WorkbookChangesPanel } from "./WorkbookChangesPanel.js";
import { useWorkbookChanges, type ZeroWorkbookChangeQuerySource } from "./use-workbook-changes.js";
import type { WorkbookChangeEntry } from "./workbook-changes-model.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function observeZeroMutationResult(result: unknown): Promise<unknown> | null {
  if (!isRecord(result)) {
    return null;
  }
  const observer = result["server"] ?? result["client"];
  return observer instanceof Promise ? observer : null;
}

export interface ZeroWorkbookChangeSource extends ZeroWorkbookChangeQuerySource {
  mutate(mutation: unknown): unknown;
}

export function useWorkbookChangesPane(input: {
  readonly documentId: string;
  readonly sheetNames: readonly string[];
  readonly zero: ZeroWorkbookChangeSource;
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
  const [pendingRevisions, setPendingRevisions] = useState<readonly number[]>([]);
  const changeCount = Math.min(changes.length, 99);

  const revertChange = useCallback(
    (change: WorkbookChangeEntry) => {
      if (!enabled || !change.canRevert || pendingRevisions.includes(change.revision)) {
        return;
      }
      setPendingRevisions((current) => [...current, change.revision]);
      const observer = observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.revertChange({
            documentId,
            revision: change.revision,
          }),
        ),
      );
      void (observer ?? Promise.resolve()).finally(() => {
        setPendingRevisions((current) =>
          current.filter((revision) => revision !== change.revision),
        );
      });
    },
    [documentId, enabled, pendingRevisions, zero],
  );

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
        onRevert={revertChange}
        pendingRevisions={pendingRevisions}
      />
    ),
    [changes, isOpen, onJump, pendingRevisions, revertChange],
  );

  return {
    changesPanel,
    changesToggle,
  };
}
