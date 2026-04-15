import { isWorkbookSnapshot, type WorkbookSnapshot } from "@bilig/protocol";
import {
  type EngineReplicaSnapshot,
  SpreadsheetEngine,
  type SpreadsheetEngineOptions,
} from "./engine.js";
import { isEngineReplicaSnapshot } from "./guards.js";

export const SPREADSHEET_ENGINE_DOCUMENT_FORMAT = "bilig.spreadsheet-engine.document.v1" as const;

export interface PersistedSpreadsheetEngineDocument {
  format: typeof SPREADSHEET_ENGINE_DOCUMENT_FORMAT;
  snapshot: WorkbookSnapshot;
  replica?: EngineReplicaSnapshot;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

/**
 * Checks whether a value matches the persisted spreadsheet engine document format.
 */
export function isPersistedSpreadsheetEngineDocument(
  value: unknown,
): value is PersistedSpreadsheetEngineDocument {
  return (
    isRecord(value) &&
    value["format"] === SPREADSHEET_ENGINE_DOCUMENT_FORMAT &&
    isWorkbookSnapshot(value["snapshot"]) &&
    (value["replica"] === undefined || isEngineReplicaSnapshot(value["replica"]))
  );
}

function assertPersistedSpreadsheetEngineDocument(
  value: unknown,
): asserts value is PersistedSpreadsheetEngineDocument {
  if (!isPersistedSpreadsheetEngineDocument(value)) {
    throw new Error("Invalid persisted spreadsheet engine document");
  }
}

/**
 * Exports a workbook snapshot and optional replica state from an engine instance.
 */
export function exportSpreadsheetEngineDocument(
  engine: SpreadsheetEngine,
  options: { includeReplica?: boolean } = {},
): PersistedSpreadsheetEngineDocument {
  const { includeReplica = true } = options;
  const document: PersistedSpreadsheetEngineDocument = {
    format: SPREADSHEET_ENGINE_DOCUMENT_FORMAT,
    snapshot: engine.exportSnapshot(),
  };
  if (includeReplica) {
    document.replica = engine.exportReplicaSnapshot();
  }
  return document;
}

/**
 * Imports a persisted workbook snapshot and optional replica state into an engine.
 */
export function importSpreadsheetEngineDocument(
  engine: SpreadsheetEngine,
  document: PersistedSpreadsheetEngineDocument,
): SpreadsheetEngine {
  assertPersistedSpreadsheetEngineDocument(document);
  engine.importSnapshot(structuredClone(document.snapshot));
  if (document.replica) {
    engine.importReplicaSnapshot(structuredClone(document.replica));
  }
  return engine;
}

/**
 * Serializes a validated spreadsheet engine document to JSON.
 */
export function serializeSpreadsheetEngineDocument(
  document: PersistedSpreadsheetEngineDocument,
): string {
  assertPersistedSpreadsheetEngineDocument(document);
  return JSON.stringify(document);
}

/**
 * Parses and validates a spreadsheet engine document from JSON.
 */
export function parseSpreadsheetEngineDocument(json: string): PersistedSpreadsheetEngineDocument {
  const parsed = JSON.parse(json) as unknown;
  assertPersistedSpreadsheetEngineDocument(parsed);
  return parsed;
}

/**
 * Creates a ready spreadsheet engine instance and hydrates it from a persisted document.
 */
export async function createSpreadsheetEngineFromDocument(
  document: PersistedSpreadsheetEngineDocument,
  options: SpreadsheetEngineOptions = {},
): Promise<SpreadsheetEngine> {
  assertPersistedSpreadsheetEngineDocument(document);
  const engine = new SpreadsheetEngine({
    ...options,
    workbookName: options.workbookName ?? document.snapshot.workbook.name,
  });
  await engine.ready();
  importSpreadsheetEngineDocument(engine, document);
  return engine;
}
