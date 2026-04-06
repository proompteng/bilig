/// <reference lib="webworker" />
import { importWorkbookFile } from "@bilig/excel-import";
import type { ImportedWorkbookPreview } from "@bilig/excel-import";
import type { WorkbookImportContentType } from "@bilig/agent-api";

declare const self: DedicatedWorkerGlobalScope;

interface WorkbookImportPreviewRequest {
  type: "preview";
  requestId: string;
  file: File;
  contentType: WorkbookImportContentType;
}

interface WorkbookImportPreviewSuccess {
  type: "success";
  requestId: string;
  preview: ImportedWorkbookPreview;
}

interface WorkbookImportPreviewError {
  type: "error";
  requestId: string;
  message: string;
}

type WorkbookImportPreviewMessage =
  | WorkbookImportPreviewRequest
  | WorkbookImportPreviewSuccess
  | WorkbookImportPreviewError;

self.addEventListener("message", (event: MessageEvent<WorkbookImportPreviewMessage>) => {
  const message = event.data;
  if (message.type !== "preview") {
    return;
  }
  void (async () => {
    try {
      const bytes = new Uint8Array(await message.file.arrayBuffer());
      const imported = importWorkbookFile(bytes, message.file.name, message.contentType);
      self.postMessage(
        {
          type: "success",
          requestId: message.requestId,
          preview: imported.preview,
        } satisfies WorkbookImportPreviewSuccess,
        [],
      );
    } catch (error) {
      self.postMessage(
        {
          type: "error",
          requestId: message.requestId,
          message: error instanceof Error ? error.message : String(error),
        } satisfies WorkbookImportPreviewError,
        [],
      );
    }
  })();
});
