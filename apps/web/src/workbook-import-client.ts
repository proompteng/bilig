import {
  CSV_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  decodeAgentFrame,
  encodeAgentFrame,
  type WorkbookImportContentType,
  type WorkbookLoadedResponse,
} from "@bilig/agent-api";
import type { ImportedWorkbookPreview } from "@bilig/excel-import";
import { resolveWorkbookNavigationUrl } from "./workbook-navigation.js";

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

type WorkbookImportPreviewResponse = WorkbookImportPreviewSuccess | WorkbookImportPreviewError;

export interface PreviewWorkbookImportInput {
  file: File;
  contentType: WorkbookImportContentType;
}

export interface FinalizeWorkbookImportInput extends PreviewWorkbookImportInput {
  openMode: "create" | "replace";
  documentId?: string;
}

function createId(prefix: string): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `${prefix}:${crypto.randomUUID()}`;
  }
  return `${prefix}:${Math.random().toString(36).slice(2)}`;
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.subarray(offset, offset + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function readErrorMessage(response: Response): Promise<string> {
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    try {
      const payload = (await response.json()) as unknown;
      if (
        typeof payload === "object" &&
        payload !== null &&
        "error" in payload &&
        typeof payload.error === "object" &&
        payload.error !== null &&
        "message" in payload.error &&
        typeof payload.error.message === "string"
      ) {
        return payload.error.message;
      }
    } catch {}
  }
  try {
    const text = await response.text();
    if (text.trim().length > 0) {
      return text;
    }
  } catch {}
  return `Workbook import failed with status ${response.status}`;
}

export function resolveWorkbookImportContentType(
  file: Pick<File, "name" | "type">,
): WorkbookImportContentType | null {
  const normalizedType = file.type.trim().toLowerCase();
  const normalizedName = file.name.trim().toLowerCase();
  if (normalizedType === XLSX_CONTENT_TYPE || normalizedName.endsWith(".xlsx")) {
    return XLSX_CONTENT_TYPE;
  }
  if (normalizedType === CSV_CONTENT_TYPE || normalizedName.endsWith(".csv")) {
    return CSV_CONTENT_TYPE;
  }
  return null;
}

export async function previewWorkbookImport(
  input: PreviewWorkbookImportInput,
): Promise<ImportedWorkbookPreview> {
  const requestId = createId("preview");
  return await new Promise<ImportedWorkbookPreview>((resolve, reject) => {
    const worker = new Worker(new URL("./workbook-import-preview.worker.ts", import.meta.url), {
      type: "module",
    });
    const cleanup = () => {
      worker.terminate();
    };
    worker.addEventListener("message", (event: MessageEvent<WorkbookImportPreviewResponse>) => {
      const payload = event.data;
      if (payload.requestId !== requestId) {
        return;
      }
      cleanup();
      if (payload.type === "error") {
        reject(new Error(payload.message));
        return;
      }
      resolve(payload.preview);
    });
    worker.addEventListener("error", (event) => {
      cleanup();
      reject(new Error(event.message || "Workbook preview worker failed"));
    });
    worker.postMessage(
      {
        type: "preview",
        requestId,
        file: input.file,
        contentType: input.contentType,
      },
      [],
    );
  });
}

export async function finalizeWorkbookImport(
  input: FinalizeWorkbookImportInput,
): Promise<WorkbookLoadedResponse> {
  const bytes = new Uint8Array(await input.file.arrayBuffer());
  const frameBytes = encodeAgentFrame({
    kind: "request",
    request: {
      kind: "loadWorkbookFile",
      id: createId("load-workbook"),
      replicaId: createId("browser-import"),
      openMode: input.openMode,
      ...(input.openMode === "replace" ? { documentId: input.documentId } : {}),
      fileName: input.file.name,
      contentType: input.contentType,
      bytesBase64: encodeBase64(bytes),
    },
  });
  const frameBody = new ArrayBuffer(frameBytes.byteLength);
  new Uint8Array(frameBody).set(frameBytes);
  const response = await fetch("/v2/agent/frames", {
    method: "POST",
    headers: {
      "content-type": "application/octet-stream",
    },
    body: frameBody,
  });

  if (!response.ok) {
    throw new Error(await readErrorMessage(response));
  }

  const frame = decodeAgentFrame(new Uint8Array(await response.arrayBuffer()));
  if (frame.kind !== "response") {
    throw new Error("Workbook import returned an invalid frame");
  }
  if (frame.response.kind === "error") {
    throw new Error(frame.response.message);
  }
  if (frame.response.kind !== "workbookLoaded") {
    throw new Error("Workbook import returned an unexpected response");
  }
  return frame.response;
}

export function resolveImportedWorkbookNavigationUrl(result: WorkbookLoadedResponse): string {
  if (result.browserUrl) {
    return result.browserUrl;
  }
  return resolveWorkbookNavigationUrl({
    documentId: result.documentId,
    serverUrl: result.serverUrl,
  });
}
