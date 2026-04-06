import { useCallback, useMemo, useState } from "react";
import type { WorkbookLoadedResponse } from "@bilig/agent-api";
import type { ImportedWorkbookPreview } from "@bilig/excel-import";
import { WorkbookImportPanel } from "./WorkbookImportPanel.js";
import {
  finalizeWorkbookImport,
  previewWorkbookImport,
  resolveImportedWorkbookNavigationUrl,
  resolveWorkbookImportContentType,
} from "./workbook-import-client.js";

interface StagedWorkbookImport {
  file: File;
  preview: ImportedWorkbookPreview;
}

export function useWorkbookImportPane(input: {
  readonly currentDocumentId: string;
  readonly enabled: boolean;
  readonly previewFile?: typeof previewWorkbookImport;
  readonly finalizeImport?: typeof finalizeWorkbookImport;
  readonly navigateToWorkbook?: (result: WorkbookLoadedResponse) => void;
}) {
  const {
    currentDocumentId,
    enabled,
    previewFile = previewWorkbookImport,
    finalizeImport = finalizeWorkbookImport,
    navigateToWorkbook = (result: WorkbookLoadedResponse) => {
      window.location.assign(resolveImportedWorkbookNavigationUrl(result));
    },
  } = input;
  const [isOpen, setIsOpen] = useState(false);
  const [isPreviewing, setIsPreviewing] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [stagedImport, setStagedImport] = useState<StagedWorkbookImport | null>(null);
  const [error, setError] = useState<string | null>(null);

  const stageFile = useCallback(
    async (file: File | null) => {
      if (!enabled || file === null) {
        return;
      }
      const contentType = resolveWorkbookImportContentType(file);
      if (!contentType) {
        setStagedImport(null);
        setError("Only local CSV and XLSX files can be staged for workbook import.");
        setIsOpen(true);
        return;
      }
      setIsOpen(true);
      setError(null);
      setIsPreviewing(true);
      try {
        const preview = await previewFile({
          file,
          contentType,
        });
        setStagedImport({
          file,
          preview,
        });
      } catch (nextError) {
        setStagedImport(null);
        setError(nextError instanceof Error ? nextError.message : String(nextError));
      } finally {
        setIsPreviewing(false);
      }
    },
    [enabled, previewFile],
  );

  const importStagedFile = useCallback(
    async (openMode: "create" | "replace") => {
      if (!enabled || !stagedImport || isImporting) {
        return;
      }
      setError(null);
      setIsImporting(true);
      try {
        const result = await finalizeImport({
          file: stagedImport.file,
          contentType: stagedImport.preview.contentType,
          openMode,
          ...(openMode === "replace" ? { documentId: currentDocumentId } : {}),
        });
        navigateToWorkbook(result);
      } catch (nextError) {
        setError(nextError instanceof Error ? nextError.message : String(nextError));
      } finally {
        setIsImporting(false);
      }
    },
    [currentDocumentId, enabled, finalizeImport, isImporting, navigateToWorkbook, stagedImport],
  );

  const importToggle = useMemo(
    () => (
      <button
        aria-controls="workbook-import-panel"
        aria-expanded={isOpen}
        aria-label="Import workbook"
        className="inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text-muted)] shadow-[var(--wb-shadow-sm)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1 disabled:cursor-not-allowed disabled:opacity-50"
        data-testid="workbook-import-toggle"
        disabled={!enabled}
        type="button"
        onClick={() => {
          setIsOpen((current) => !current);
        }}
      >
        Import
      </button>
    ),
    [enabled, isOpen],
  );

  const importPanel = useMemo(
    () => (
      <WorkbookImportPanel
        currentDocumentId={currentDocumentId}
        enabled={enabled}
        error={error}
        isImporting={isImporting}
        isOpen={isOpen}
        isPreviewing={isPreviewing}
        stagedPreview={stagedImport?.preview ?? null}
        onClose={() => {
          setIsOpen(false);
        }}
        onFileSelected={(file) => {
          void stageFile(file);
        }}
        onImportAsNew={() => {
          void importStagedFile("create");
        }}
        onReplaceCurrent={() => {
          void importStagedFile("replace");
        }}
      />
    ),
    [
      currentDocumentId,
      enabled,
      error,
      importStagedFile,
      isImporting,
      isOpen,
      isPreviewing,
      stageFile,
      stagedImport?.preview,
    ],
  );

  return {
    importPanel,
    importToggle,
  };
}
