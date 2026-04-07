import { useCallback, useMemo, useState } from "react";
import { mutators } from "@bilig/zero-sync";
import type { Viewport } from "@bilig/protocol";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import { WorkbookViewsPanel } from "./WorkbookViewsPanel.js";
import {
  WorkbookHeaderActionButton,
  WorkbookHeaderCountBadge,
} from "./workbook-header-controls.js";
import { useWorkbookViews, type ZeroWorkbookSheetViewQuerySource } from "./use-workbook-views.js";
import type {
  WorkbookSheetViewEntry,
  WorkbookSheetViewVisibility,
} from "./workbook-views-model.js";

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

function createWorkbookViewId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `view:${Math.random().toString(36).slice(2)}`;
}

function defaultWorkbookViewName(selection: WorkerRuntimeSelection): string {
  return `${selection.sheetName} ${selection.address}`;
}

export interface ZeroWorkbookSheetViewSource extends ZeroWorkbookSheetViewQuerySource {
  mutate(mutation: unknown): unknown;
}

export function buildWorkbookSheetViewMutationArgs(input: {
  readonly documentId: string;
  readonly id: string;
  readonly name: string;
  readonly visibility: WorkbookSheetViewVisibility;
  readonly selection: WorkerRuntimeSelection;
  readonly viewport: Viewport;
}) {
  return {
    documentId: input.documentId,
    id: input.id,
    name: input.name,
    visibility: input.visibility,
    sheetName: input.selection.sheetName,
    address: input.selection.address,
    viewport: input.viewport,
  };
}

export function useWorkbookViewsPane(input: {
  readonly documentId: string;
  readonly currentUserId: string;
  readonly selection: WorkerRuntimeSelection;
  readonly sheetNames: readonly string[];
  readonly zero: ZeroWorkbookSheetViewSource;
  readonly enabled: boolean;
  readonly getCurrentViewport: () => Viewport;
  readonly onApply: (view: WorkbookSheetViewEntry) => void;
}) {
  const {
    currentUserId,
    documentId,
    enabled,
    getCurrentViewport,
    onApply,
    selection,
    sheetNames,
    zero,
  } = input;
  const views = useWorkbookViews({
    currentUserId,
    documentId,
    sheetNames,
    zero,
    enabled,
  });
  const [isOpen, setIsOpen] = useState(false);
  const [draftName, setDraftName] = useState("");
  const [draftVisibility, setDraftVisibility] = useState<WorkbookSheetViewVisibility>("private");
  const viewCount = Math.min(views.length, 99);

  const saveView = useCallback(
    (existing?: Pick<WorkbookSheetViewEntry, "id" | "name" | "visibility">) => {
      if (!enabled) {
        return;
      }
      const resolvedName = draftName.trim() || existing?.name || defaultWorkbookViewName(selection);
      const args = buildWorkbookSheetViewMutationArgs({
        documentId,
        id: existing?.id ?? createWorkbookViewId(),
        name: resolvedName,
        visibility: existing?.visibility ?? draftVisibility,
        selection,
        viewport: getCurrentViewport(),
      });
      observeZeroMutationResult(zero.mutate(mutators.workbook.upsertSheetView(args)));
      if (!existing) {
        setDraftName("");
      }
    },
    [documentId, draftName, draftVisibility, enabled, getCurrentViewport, selection, zero],
  );

  const deleteView = useCallback(
    (view: WorkbookSheetViewEntry) => {
      if (!enabled) {
        return;
      }
      observeZeroMutationResult(
        zero.mutate(
          mutators.workbook.deleteSheetView({
            documentId,
            id: view.id,
          }),
        ),
      );
    },
    [documentId, enabled, zero],
  );

  const viewsToggle = useMemo(
    () => (
      <WorkbookHeaderActionButton
        aria-controls="workbook-views-panel"
        aria-expanded={isOpen}
        aria-label={`Show workbook views (${views.length})`}
        data-testid="workbook-views-toggle"
        isActive={isOpen}
        isGrouped
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        <span>Views</span>
        <WorkbookHeaderCountBadge value={viewCount} />
      </WorkbookHeaderActionButton>
    ),
    [isOpen, viewCount, views.length],
  );

  const viewsPanel = useMemo(
    () => (
      <WorkbookViewsPanel
        draftName={draftName}
        draftVisibility={draftVisibility}
        isOpen={isOpen}
        views={views}
        onApply={(view) => {
          onApply(view);
        }}
        onClose={() => {
          setIsOpen(false);
        }}
        onDelete={deleteView}
        onDraftNameChange={setDraftName}
        onDraftVisibilityChange={setDraftVisibility}
        onSave={() => {
          saveView();
        }}
        onUpdate={(view) => {
          saveView(view);
        }}
      />
    ),
    [deleteView, draftName, draftVisibility, isOpen, onApply, saveView, views],
  );

  return {
    viewsPanel,
    viewsToggle,
  };
}
