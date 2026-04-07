import { useId } from "react";
import { Upload } from "lucide-react";
import type { WorkbookImportContentType } from "@bilig/agent-api";
import type { ImportedWorkbookPreview } from "@bilig/excel-import";
import { cn } from "./cn.js";
import {
  workbookAlertClass,
  workbookButtonClass,
  workbookPillClass,
  workbookSurfaceClass,
} from "./workbook-shell-chrome.js";

function formatFileSize(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatImportType(contentType: WorkbookImportContentType): string {
  return contentType === "text/csv" ? "CSV" : "XLSX";
}

function createPreviewRowDescriptors(rows: readonly (readonly string[])[]): readonly {
  key: string;
  cells: readonly {
    key: string;
    value: string;
  }[];
}[] {
  const rowCounts = new Map<string, number>();
  return rows.map((row) => {
    const rowBaseKey = JSON.stringify(row);
    const rowOccurrence = (rowCounts.get(rowBaseKey) ?? 0) + 1;
    rowCounts.set(rowBaseKey, rowOccurrence);

    const cellCounts = new Map<string, number>();
    const cells = row.map((value) => {
      const cellBaseKey = value || "__blank__";
      const cellOccurrence = (cellCounts.get(cellBaseKey) ?? 0) + 1;
      cellCounts.set(cellBaseKey, cellOccurrence);
      return {
        key: `${cellBaseKey}:${cellOccurrence}`,
        value,
      };
    });

    return {
      key: `${rowBaseKey}:${rowOccurrence}`,
      cells,
    };
  });
}

function WorkbookImportSheetPreview(props: {
  readonly preview: ImportedWorkbookPreview["sheets"][number];
}) {
  const rows = createPreviewRowDescriptors(props.preview.previewRows);

  return (
    <section className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-4">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="text-[13px] font-semibold text-[var(--wb-text)]">
            {props.preview.name}
          </div>
          <div className="mt-1 text-[11px] text-[var(--wb-text-subtle)]">
            {props.preview.rowCount} rows · {props.preview.columnCount} columns ·{" "}
            {props.preview.nonEmptyCellCount} populated cells
          </div>
        </div>
      </div>
      {props.preview.previewRows.length > 0 ? (
        <div className="mt-3 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
          <table className="min-w-full border-collapse text-left text-[11px]">
            <tbody>
              {rows.map((row) => (
                <tr
                  key={`${props.preview.name}:${row.key}`}
                  className="border-t border-[var(--wb-border)] first:border-t-0"
                >
                  {row.cells.map((cell) => (
                    <td
                      key={`${props.preview.name}:${row.key}:${cell.key}`}
                      className="max-w-[10rem] truncate border-l border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-2 py-1.5 align-top first:border-l-0"
                      title={cell.value}
                    >
                      {cell.value || <span className="text-[var(--wb-text-subtle)]">(blank)</span>}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="mt-3 rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-4 text-[12px] text-[var(--wb-text-subtle)]">
          Empty
        </div>
      )}
    </section>
  );
}

export function WorkbookImportPanel(props: {
  readonly isOpen: boolean;
  readonly enabled: boolean;
  readonly stagedPreview: ImportedWorkbookPreview | null;
  readonly isPreviewing: boolean;
  readonly isImporting: boolean;
  readonly onClose: () => void;
  readonly onFileSelected: (file: File | null) => void;
  readonly onImportAsNew: () => void;
  readonly onReplaceCurrent: () => void;
}) {
  const fileInputId = useId();

  if (!props.isOpen) {
    return null;
  }

  return (
    <div
      className="absolute inset-0 z-30 flex items-center justify-center bg-[rgba(28,25,38,0.14)] p-4"
      data-testid="workbook-import-panel"
      id="workbook-import-panel"
    >
      <div
        aria-label="Workbook import staging"
        aria-modal="true"
        className="relative flex max-h-[calc(100vh-3rem)] w-full max-w-[72rem] flex-col overflow-hidden rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] shadow-[0_12px_32px_rgba(28,25,38,0.1)]"
        role="dialog"
      >
        <button
          aria-label="Close workbook import"
          className="absolute top-4 right-4 inline-flex h-8 w-8 items-center justify-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] text-[var(--wb-text-muted)] transition-colors hover:text-[var(--wb-text)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
          type="button"
          onClick={props.onClose}
        >
          ×
        </button>
        <div className="grid min-h-0 flex-1 gap-0 lg:grid-cols-[18rem,minmax(0,1fr)]">
          <div className="flex flex-col gap-3 border-b border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] p-4 pr-12 lg:border-b-0 lg:border-r">
            <label className="flex flex-col gap-3">
              <span className="sr-only">Select file</span>
              <div
                className="rounded-[var(--wb-radius-control)] border border-dashed border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 py-3"
                data-testid="workbook-import-picker-shell"
              >
                <div className="flex items-center gap-3 text-[12px] text-[var(--wb-text-muted)]">
                  <Upload className="h-4 w-4" />
                  <span className="truncate">{props.stagedPreview?.fileName ?? "File"}</span>
                </div>
                <input
                  accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  className="sr-only"
                  data-testid="workbook-import-file"
                  disabled={!props.enabled || props.isPreviewing || props.isImporting}
                  id={fileInputId}
                  type="file"
                  onChange={(event) => {
                    props.onFileSelected(event.currentTarget.files?.[0] ?? null);
                    event.currentTarget.value = "";
                  }}
                />
                <div className="mt-3 flex items-center gap-2">
                  <label className={workbookButtonClass({ tone: "neutral" })} htmlFor={fileInputId}>
                    Choose
                  </label>
                  {props.stagedPreview?.contentType ? (
                    <span className="min-w-0 truncate text-[11px] text-[var(--wb-text-subtle)]">
                      {formatImportType(props.stagedPreview.contentType)}
                    </span>
                  ) : null}
                </div>
              </div>
            </label>

            {props.isPreviewing ? (
              <div
                className={cn(
                  workbookSurfaceClass(),
                  "px-3 py-3 text-[12px] text-[var(--wb-text-muted)]",
                )}
              >
                Loading…
              </div>
            ) : null}

            {props.stagedPreview ? (
              <div className="flex flex-col gap-3">
                <div className="rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-4 py-4">
                  <div className="flex items-center gap-2">
                    <span className={workbookPillClass({ tone: "accent", weight: "strong" })}>
                      {formatImportType(props.stagedPreview.contentType)}
                    </span>
                    <span className="truncate text-[13px] font-semibold text-[var(--wb-text)]">
                      {props.stagedPreview.fileName}
                    </span>
                  </div>
                  <div className="mt-3 grid gap-1 text-[12px] text-[var(--wb-text-muted)]">
                    <div className="text-[13px] text-[var(--wb-text)]">
                      {props.stagedPreview.workbookName}
                    </div>
                    <div>
                      {props.stagedPreview.sheetCount} sheets ·{" "}
                      {formatFileSize(props.stagedPreview.fileSizeBytes)}
                    </div>
                  </div>
                </div>

                {props.stagedPreview.warnings.length > 0 ? (
                  <div className={cn(workbookAlertClass({ tone: "warning" }), "px-4 py-3")}>
                    <ul className="list-disc space-y-1 pl-4">
                      {props.stagedPreview.warnings.map((warning) => (
                        <li key={warning}>{warning}</li>
                      ))}
                    </ul>
                  </div>
                ) : null}

                <div className="mt-auto flex gap-2">
                  <button
                    className={workbookButtonClass({
                      tone: "accent",
                      size: "md",
                      weight: "strong",
                    })}
                    data-testid="workbook-import-create"
                    disabled={props.isImporting}
                    style={{ flex: 1 }}
                    type="button"
                    onClick={props.onImportAsNew}
                  >
                    {props.isImporting ? "…" : "New"}
                  </button>
                  <button
                    className={workbookButtonClass({
                      tone: "neutral",
                      size: "md",
                    })}
                    data-testid="workbook-import-replace"
                    disabled={props.isImporting}
                    style={{ flex: 1 }}
                    type="button"
                    onClick={props.onReplaceCurrent}
                  >
                    Replace
                  </button>
                </div>
              </div>
            ) : null}
          </div>

          <div className="min-h-0 overflow-y-auto p-5">
            {props.stagedPreview ? (
              <div className="flex flex-col gap-4">
                {props.stagedPreview.sheets.map((sheet) => (
                  <WorkbookImportSheetPreview key={sheet.name} preview={sheet} />
                ))}
              </div>
            ) : (
              <div />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
