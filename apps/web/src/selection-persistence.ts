import type { WorkerRuntimeSelection } from "./runtime-session.js";

const DEFAULT_SELECTION: WorkerRuntimeSelection = {
  sheetName: "Sheet1",
  address: "A1",
};

function storageKey(documentId: string): string {
  return `bilig:selection:${documentId}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function loadPersistedSelection(documentId: string): WorkerRuntimeSelection {
  if (typeof window === "undefined") {
    return DEFAULT_SELECTION;
  }
  try {
    const raw = window.localStorage.getItem(storageKey(documentId));
    if (!raw) {
      return DEFAULT_SELECTION;
    }
    const parsed = JSON.parse(raw);
    if (
      !isRecord(parsed) ||
      typeof parsed["sheetName"] !== "string" ||
      parsed["sheetName"].trim().length === 0 ||
      typeof parsed["address"] !== "string" ||
      parsed["address"].trim().length === 0
    ) {
      return DEFAULT_SELECTION;
    }
    return {
      sheetName: parsed["sheetName"],
      address: parsed["address"],
    };
  } catch {
    return DEFAULT_SELECTION;
  }
}

export function persistSelection(documentId: string, selection: WorkerRuntimeSelection): void {
  if (typeof window === "undefined") {
    return;
  }
  try {
    window.localStorage.setItem(storageKey(documentId), JSON.stringify(selection));
  } catch {
    // Ignore storage failures and keep the runtime usable.
  }
}
