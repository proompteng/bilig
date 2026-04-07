import { useCallback, useMemo, useState } from "react";
import { mutators } from "@bilig/zero-sync";
import type { Viewport } from "@bilig/protocol";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import { WorkbookVersionsPanel } from "./WorkbookVersionsPanel.js";
import {
  WorkbookHeaderActionButton,
  WorkbookHeaderCountBadge,
} from "./workbook-header-controls.js";
import {
  useWorkbookVersions,
  type ZeroWorkbookVersionQuerySource,
} from "./use-workbook-versions.js";
import type { WorkbookVersionEntry } from "./workbook-versions-model.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function observeZeroMutationResult(result: unknown): void {
  if (!isRecord(result)) {
    return;
  }
  const observer = result["server"] ?? result["client"];
  if (!(observer instanceof Promise)) {
    return;
  }
  void observer.catch(() => undefined);
}

function createWorkbookVersionId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `version:${Math.random().toString(36).slice(2)}`;
}

function defaultWorkbookVersionName(selection: WorkerRuntimeSelection): string {
  return `${selection.sheetName} ${selection.address}`;
}

export interface ZeroWorkbookVersionSource extends ZeroWorkbookVersionQuerySource {
  mutate(mutation: unknown): unknown;
}

export function buildWorkbookVersionMutationArgs(input: {
  readonly documentId: string;
  readonly id: string;
  readonly name: string;
  readonly selection: WorkerRuntimeSelection;
  readonly viewport: Viewport;
}) {
  return {
    documentId: input.documentId,
    id: input.id,
    name: input.name,
    sheetName: input.selection.sheetName,
    address: input.selection.address,
    viewport: input.viewport,
  };
}

export function useWorkbookVersionsPane(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly selection: WorkerRuntimeSelection;
  readonly zero: ZeroWorkbookVersionSource;
  readonly enabled: boolean;
  readonly getCurrentViewport: () => Viewport;
}) {
  const { currentUserId, documentId, enabled, getCurrentViewport, selection, zero } = input;
  const versions = useWorkbookVersions({
    documentId,
    currentUserId,
    zero,
    enabled,
  });
  const [isOpen, setIsOpen] = useState(false);
  const [draftName, setDraftName] = useState("");
  const versionCount = Math.min(versions.length, 99);

  const saveVersion = useCallback(() => {
    if (!enabled) {
      return;
    }
    const resolvedName = draftName.trim() || defaultWorkbookVersionName(selection);
    observeZeroMutationResult(
      zero.mutate(
        mutators.workbook.createVersion(
          buildWorkbookVersionMutationArgs({
            documentId,
            id: createWorkbookVersionId(),
            name: resolvedName,
            selection,
            viewport: getCurrentViewport(),
          }),
        ),
      ),
    );
    setDraftName("");
  }, [documentId, draftName, enabled, getCurrentViewport, selection, zero]);

  const deleteVersion = useCallback(
    (version: WorkbookVersionEntry) => {
      if (!enabled) {
        return;
      }
      observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.deleteVersion({
            documentId,
            id: version.id,
          }),
        ),
      );
    },
    [documentId, enabled, zero],
  );

  const restoreVersion = useCallback(
    (version: WorkbookVersionEntry) => {
      if (!enabled) {
        return;
      }
      observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.restoreVersion({
            documentId,
            id: version.id,
          }),
        ),
      );
    },
    [documentId, enabled, zero],
  );

  const versionsToggle = useMemo(
    () => (
      <WorkbookHeaderActionButton
        aria-controls="workbook-versions-panel"
        aria-expanded={isOpen}
        aria-label={`Show workbook versions (${versions.length})`}
        data-testid="workbook-versions-toggle"
        isActive={isOpen}
        isGrouped
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        <span>Versions</span>
        <WorkbookHeaderCountBadge value={versionCount} />
      </WorkbookHeaderActionButton>
    ),
    [isOpen, versionCount, versions.length],
  );

  const versionsPanel = useMemo(
    () => (
      <WorkbookVersionsPanel
        draftName={draftName}
        isOpen={isOpen}
        versions={versions}
        onClose={() => {
          setIsOpen(false);
        }}
        onDelete={deleteVersion}
        onDraftNameChange={setDraftName}
        onRestore={restoreVersion}
        onSave={saveVersion}
      />
    ),
    [deleteVersion, draftName, isOpen, restoreVersion, saveVersion, versions],
  );

  return {
    versionsPanel,
    versionsToggle,
  };
}
